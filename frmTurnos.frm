VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmTurnos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Datos de turnos"
   ClientHeight    =   5925
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10695
   Icon            =   "frmTurnos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5925
   ScaleWidth      =   10695
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnBuscar 
      Appearance      =   0  'Flat
      Caption         =   "+"
      Height          =   290
      Index           =   2
      Left            =   900
      MaskColor       =   &H00000000&
      TabIndex        =   28
      ToolTipText     =   "Buscar Fecha"
      Top             =   4920
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Frame Frame2 
      Caption         =   "Datos del Turno"
      ForeColor       =   &H00972E0B&
      Height          =   4410
      Left            =   5880
      TabIndex        =   19
      Top             =   600
      Width           =   4695
      Begin VB.TextBox txtAux 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   5
         Left            =   120
         MaxLength       =   6
         TabIndex        =   6
         Tag             =   "Artículo|N|S|0|999999|sturno|codartic|000000||"
         Top             =   840
         Width           =   735
      End
      Begin VB.TextBox txtAux2 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   5
         Left            =   960
         TabIndex        =   29
         Top             =   840
         Width           =   3495
      End
      Begin VB.TextBox txtAux 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   9
         Left            =   120
         MaxLength       =   12
         TabIndex        =   10
         Tag             =   "Contador inicio|N|S|0|99999999.99|sturno|containi|##,###,##0.00||"
         Top             =   2880
         Width           =   1215
      End
      Begin VB.TextBox txtAux 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   10
         Left            =   2160
         MaxLength       =   12
         TabIndex        =   11
         Tag             =   "Contador final|N|S|0|99999999.99|sturno|contafin|##,###,##0.00||"
         Top             =   2880
         Width           =   1215
      End
      Begin VB.TextBox txtAux 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   6
         Left            =   120
         MaxLength       =   1
         TabIndex        =   7
         Tag             =   "Tipo Crédito|N|S|0|9|sturno|tipocred|0||"
         Top             =   1560
         Width           =   735
      End
      Begin VB.TextBox txtAux2 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   6
         Left            =   960
         TabIndex        =   20
         Top             =   1560
         Width           =   3495
      End
      Begin VB.TextBox txtAux 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   8
         Left            =   2160
         MaxLength       =   12
         TabIndex        =   9
         Tag             =   "Importe|N|S|0|99999999.99|sturno|importel|###,##0.00||"
         Top             =   2280
         Width           =   1215
      End
      Begin VB.TextBox txtAux 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   7
         Left            =   120
         MaxLength       =   12
         TabIndex        =   8
         Tag             =   "Litros|N|S|0|99999999.99|sturno|litrosve|###,##0.00||"
         Top             =   2280
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Contador inicial"
         Height          =   255
         Index           =   9
         Left            =   120
         TabIndex        =   26
         Top             =   2640
         Width           =   1215
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   0
         Left            =   960
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar Artículo"
         Top             =   600
         Width           =   240
      End
      Begin VB.Label Label2 
         Caption         =   "Contador final"
         Height          =   255
         Left            =   2160
         TabIndex        =   25
         Top             =   2640
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Artículo"
         Height          =   255
         Index           =   7
         Left            =   120
         TabIndex        =   24
         Top             =   600
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Tipo Crédito"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   23
         Top             =   1320
         Width           =   855
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   1
         Left            =   1080
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar Tipo Crédito"
         Top             =   1320
         Width           =   240
      End
      Begin VB.Label Label4 
         Caption         =   "Importe"
         Height          =   255
         Left            =   2160
         TabIndex        =   22
         Top             =   2040
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Litros"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   21
         Top             =   2040
         Width           =   1215
      End
   End
   Begin VB.TextBox txtAux 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   290
      Index           =   4
      Left            =   4440
      MaxLength       =   3
      TabIndex        =   5
      Tag             =   "Manguera|N|S|0|999|sturno|nummangu|000||"
      Text            =   "mang"
      Top             =   4920
      Width           =   555
   End
   Begin VB.ComboBox Combo1 
      Appearance      =   0  'Flat
      Height          =   315
      ItemData        =   "frmTurnos.frx":000C
      Left            =   2400
      List            =   "frmTurnos.frx":000E
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Tag             =   "Tipo Registro|N|N|0|4|sturno|tiporegi|||"
      Top             =   4920
      Width           =   1215
   End
   Begin VB.TextBox txtAux 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   290
      Index           =   2
      Left            =   1800
      MaxLength       =   4
      TabIndex        =   2
      Tag             =   "Linea|N|N|0|9999|sturno|numlinea|0000|S|"
      Text            =   "linea"
      Top             =   4920
      Width           =   555
   End
   Begin VB.TextBox txtAux 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   290
      Index           =   0
      Left            =   120
      MaxLength       =   10
      TabIndex        =   0
      Tag             =   "Fecha|F|N|||sturno|fechatur|dd/mm/yyyy|S|"
      Text            =   "fec"
      Top             =   4920
      Width           =   795
   End
   Begin VB.TextBox txtAux 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   290
      Index           =   3
      Left            =   3720
      MaxLength       =   3
      TabIndex        =   4
      Tag             =   "Tanque|N|S|0|999|sturno|numtanqu|000||"
      Text            =   "tanq"
      Top             =   4920
      Width           =   555
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   7860
      TabIndex        =   12
      Top             =   5340
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   9120
      TabIndex        =   13
      Top             =   5340
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox txtAux 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   290
      Index           =   1
      Left            =   960
      MaxLength       =   1
      TabIndex        =   1
      Tag             =   "Turno|N|N|1|9|sturno|codturno|0|S|"
      Text            =   "turno"
      Top             =   4920
      Width           =   795
   End
   Begin VB.CommandButton cmdRegresar 
      Caption         =   "&Regresar"
      Height          =   375
      Left            =   9120
      TabIndex        =   18
      Top             =   5340
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Height          =   555
      Index           =   1
      Left            =   120
      TabIndex        =   14
      Top             =   5175
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
         Left            =   40
         TabIndex        =   15
         Top             =   240
         Width           =   2295
      End
   End
   Begin MSAdodcLib.Adodc adodc1 
      Height          =   330
      Left            =   4440
      Top             =   120
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
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
      TabIndex        =   16
      Top             =   0
      Width           =   10695
      _ExtentX        =   18865
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   12
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Buscar"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Ver todos"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
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
            Object.ToolTipText     =   "Generación Masiva"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Imprimir"
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin VB.CheckBox chkVistaPrevia 
         Caption         =   "Vista previa"
         Height          =   195
         Left            =   8400
         TabIndex        =   17
         Top             =   120
         Visible         =   0   'False
         Width           =   1215
      End
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmTurnos.frx":0010
      Height          =   4410
      Left            =   120
      TabIndex        =   27
      Top             =   540
      Width           =   5595
      _ExtentX        =   9869
      _ExtentY        =   7779
      _Version        =   393216
      AllowUpdate     =   0   'False
      BorderStyle     =   0
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         AllowFocus      =   0   'False
         AllowRowSizing  =   0   'False
         AllowSizing     =   0   'False
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnOpciones 
      Caption         =   "&Opciones"
      Begin VB.Menu mnBuscar 
         Caption         =   "&Buscar"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnVerTodos 
         Caption         =   "&Ver todos"
         Shortcut        =   ^B
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
      Begin VB.Menu mnGeneracion 
         Caption         =   "Generación Masiva"
         Shortcut        =   ^G
      End
      Begin VB.Menu mnBarra2 
         Caption         =   "-"
      End
      Begin VB.Menu mnSalir 
         Caption         =   "&Salir"
         Shortcut        =   ^S
      End
   End
End
Attribute VB_Name = "frmTurnos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' +-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+
' +-+- Autor: MANOLO                               -+-+
' +-+- Menú: DATOS DE TURNOS                       -+-+
' +-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+

' +-+-+-+- DISSENY +-+-+-+-
' 1. Posar tots els controls al formulari
' 2. Posar els index correlativament
' 3. Si n'hi han botons de buscar repasar el ToolTipText
' 4. Alliniar els camps numérics a la dreta i el resto a l'esquerra
' 5. Posar els TAGs
' (si es INTEGER: si PK => mínim 1; si no PK => mínim 0; màxim => 99; format => 00)
' (si es DECIMAL; mínim => 0; màxim => 99.99; format => #,###,###,##0.00)
' (si es DATE; format => dd/mm/yyyy)
' 6. Posar els MAXLENGTHs
' 7. Posar els TABINDEXs

Option Explicit

Public DatosADevolverBusqueda As String    'Tendra el nº de text que quiere que devuelva, empipados
Public Event DatoSeleccionado(CadenaSeleccion As String)

'codi per al registe que s'afegix al cridar des d'altre formulari.
'Obrir en modo Insertar i tornar datos del registre insertat
Public NuevoCodigo As String

Public DeConsulta As Boolean

' *** declarar els formularis als que vaig a cridar ***
'Private WithEvents frmB As frmBuscaGrid
Private WithEvents frmArt As frmManArtic
Attribute frmArt.VB_VarHelpID = -1
Private WithEvents frmTCred As frmManTipCre
Attribute frmTCred.VB_VarHelpID = -1
Private WithEvents frmC As frmCal
Attribute frmC.VB_VarHelpID = -1

Private CadenaConsulta As String
Private CadB As String

Dim Modo As Byte
'----------- MODOS --------------------------------
'   0.-  Formulari net sense cap camp ple
'   1.-  Preparant per a fer la búsqueda
'   2.-  Ja tenim registres i els anem a recorrer
'        ,podem editar-los Edició del camp
'   3.-  Inserció de nou registre
'   4.-  Modificar
'--------------------------------------------------


Private Sub PonerModo(vModo)
Dim b As Boolean
Dim i As Integer

    Modo = vModo
'    PonerIndicador lblIndicador, Modo

    b = (Modo = 2)
    If b Then
        PonerContRegIndicador
    Else
        PonerIndicador lblIndicador, Modo
    End If
    
    ' **** posar tots els controls (botons inclosos) que siguen del Grid
    txtAux(0).visible = Not b
    txtAux(1).visible = Not b
    txtAux(2).visible = False ' Linea
    Combo1.visible = Not b
    txtAux(3).visible = Not b
    txtAux(4).visible = Not b
    btnBuscar(2).visible = Not b

    ' **** si n'hi han controls (imagens incloses) fora del grid, bloquejar-los;
    ' no posar els camps de descripció de fora del grid ****
    For i = 5 To 10
        BloquearTxt txtAux(i), b
    Next i

    BloquearImgBuscar Me, Modo  ' ** si n'hi han imagens de buscar codi fora del grid **
    ' ********************************************************

    cmdAceptar.visible = Not b
    cmdCancelar.visible = Not b
    DataGrid1.Enabled = b

    'Si es retornar
    If DatosADevolverBusqueda <> "" Then cmdRegresar.visible = b

    PonerLongCampos
    PonerModoOpcionesMenu 'Activar/Desact botons de menu según Modo
    PonerOpcionesMenu 'Activar/Desact botons de menu según permissos de l'usuari

    ' *** bloquejar tota la PK quan estem en modificar  ***
    BloquearTxt txtAux(0), (Modo = 4) 'fecha
    BloquearTxt txtAux(1), (Modo = 4) 'turno
    BloquearTxt txtAux(2), (Modo = 4) 'linea
    BloquearBtn btnBuscar(2), (Modo = 4) 'boton de la fecha
    ' ******************************************************
End Sub
Private Sub PonerModoOpcionesMenu()
'Activa/Desactiva botones del la toobar y del menu, segun el modo en que estemos
Dim b As Boolean

    b = (Modo = 2)
    'Busqueda
    Toolbar1.Buttons(2).Enabled = b
    Me.mnBuscar.Enabled = b
    'Ver Todos
    Toolbar1.Buttons(3).Enabled = b
    Me.mnVerTodos.Enabled = b
    
    'Insertar
    Toolbar1.Buttons(6).Enabled = b And Not DeConsulta
    Me.mnNuevo.Enabled = b And Not DeConsulta
    
    b = (b And adodc1.Recordset.RecordCount > 0) And Not DeConsulta
    'Modificar
    Toolbar1.Buttons(7).Enabled = b
    Me.mnModificar.Enabled = b
    'Eliminar
    Toolbar1.Buttons(8).Enabled = b
    Me.mnEliminar.Enabled = b
    'Imprimir
    Toolbar1.Buttons(11).Enabled = b
'    Me.mnImprimir.Enabled = b
End Sub
'
'Private Sub PonerModoOpcionesMenu()
''Activa/Desactiva botons de la toolbar i del menu, según el modo en que estiguem
'Dim b As Boolean
'
'    ' *** adrede: per a que no es puga fer res si estic cridant des de frmViagrc ***
'
'    b = (Modo = 2)
'    'Busqueda
'    Toolbar1.Buttons(2).Enabled = b And ExpedBusca = 0
'    Me.mnBuscar.Enabled = b And ExpedBusca = 0
'    'Vore Tots
'    Toolbar1.Buttons(3).Enabled = b And ExpedBusca = 0
'    Me.mnVerTodos.Enabled = b And ExpedBusca = 0
'
'    'Insertar
'    Toolbar1.Buttons(6).Enabled = b And Not DeConsulta And ExpedBusca = 0
'    Me.mnNuevo.Enabled = b And Not DeConsulta And ExpedBusca = 0
'
'    b = (b And adodc1.Recordset.RecordCount > 0) And Not DeConsulta And ExpedBusca = 0
'    'Modificar
'    Toolbar1.Buttons(7).Enabled = b And ExpedBusca = 0
'    Me.mnModificar.Enabled = b And ExpedBusca = 0
'    'Eliminar
'    Toolbar1.Buttons(8).Enabled = b And ExpedBusca = 0
'    Me.mnEliminar.Enabled = b And ExpedBusca = 0
'    'Imprimir
'    Toolbar1.Buttons(11).Enabled = b And ExpedBusca = 0
'
'    ' ******************************************************************************
'End Sub

Private Sub BotonAnyadir()
Dim NumF As String
Dim anc As Single
Dim i As Integer

    CargaGrid 'primer de tot carregue tot el grid
    CadB = ""
    '********* canviar taula i camp; repasar codEmpre ************
    If (DatosADevolverBusqueda <> "") And NuevoCodigo <> "" Then
        NumF = NuevoCodigo
    Else
        'NumF = SugerirCodigoSiguienteStr("follviaj", "codfovia")
        'NumF = SugerirCodigoSiguienteStr("sturno", "numlinea", "codempre=" & vSesion.Empresa)
    End If
    '***************************************************************
    'Situem el grid al final
    AnyadirLinea DataGrid1, adodc1

    anc = DataGrid1.Top
    If DataGrid1.Row < 0 Then
        anc = anc + 206
    Else
        anc = anc + DataGrid1.RowTop(DataGrid1.Row) + 5
    End If

    ' *** valors per defecte a l'afegir (dins i fora del grid); repasar codEmpre ***
    txtAux(0).Text = Format(Now, "dd/mm/yyyy")
    txtAux(1).Text = ""
    'FormateaCampo txtAux(1)
    For i = 3 To 10
        txtAux(i).Text = ""
    Next i
    For i = 5 To 6
        txtAux2(i).Text = ""
    Next i
    Combo1.ListIndex = -1 'per defecte a null
    ' **************************************************

    LLamaLineas anc, 3

    ' *** posar el foco ***
    If (DatosADevolverBusqueda <> "") And NuevoCodigo <> "" Then
        PonerFoco txtAux(3) '**** 1r camp visible del grid que NO siga PK ****
    Else
        PonerFoco txtAux(0) '**** 1r camp visible del grid que siga PK ****
    End If
    ' ******************************************************
End Sub

Private Sub BotonVerTodos()
    CargaGrid ""
    CadB = ""
    PonerModo 2
End Sub

Private Sub BotonBuscar()
    Dim i As Integer
    ' *** canviar per la PK (no posar codempre si està a Form_Load) ***
    'CargaGrid "codsupdt = -1 AND codempre = " & codEmpre
    CargaGrid "codturno = -1"
    '*******************************************************************************

    ' *** canviar-ho pels valors per defecte al buscar (dins i fora del grid);
    ' repasar codEmpre ******
    For i = 0 To 10
        txtAux(i).Text = ""
    Next i
    For i = 5 To 6
        txtAux2(i).Text = ""
    Next i
    Combo1.ListIndex = -1
    ' ****************************************************

    LLamaLineas DataGrid1.Top + 206, 1

    ' *** posar el foco al 1r camp visible del grid que siga PK ***
    PonerFoco txtAux(0)
    ' ***************************************************************
End Sub

Private Sub BotonModificar()
    Dim anc As Single
    Dim i As Integer

    Screen.MousePointer = vbHourglass

    If DataGrid1.Bookmark < DataGrid1.FirstRow Or DataGrid1.Bookmark > (DataGrid1.FirstRow + DataGrid1.VisibleRows - 1) Then
        i = DataGrid1.Bookmark - DataGrid1.FirstRow
        DataGrid1.Scroll 0, i
        DataGrid1.Refresh
    End If

    If DataGrid1.Row < 0 Then
        anc = 320
    Else
        anc = DataGrid1.RowTop(DataGrid1.Row) + 545
    End If

    ' *** asignar als controls del grid, els valors de les columnes ***
    txtAux(0).Text = DataGrid1.Columns(0).Text
    txtAux(1).Text = DataGrid1.Columns(1).Text
    txtAux(2).Text = DataGrid1.Columns(2).Text
    i = adodc1.Recordset!tiporegi
    PosicionarCombo Me.Combo1, i
    txtAux(3).Text = DataGrid1.Columns(5).Text
    txtAux(4).Text = DataGrid1.Columns(6).Text
    ' ********************************************************

    LLamaLineas anc, 4 'modo 4

    ' *** foco al 1r camp visible que NO siga clau primaria ***
    PonerFocoCmb Combo1
    ' *********************************************************
    Screen.MousePointer = vbDefault
End Sub

Private Sub LLamaLineas(alto As Single, xModo As Byte)
Dim i As Integer

    DeseleccionaGrid Me.DataGrid1
    PonerModo xModo

    ' *** posar el Top a tots els controls del grid (botons també) ***
    'Me.imgFec(2).Top = alto
    For i = 0 To 4
        txtAux(i).Top = alto
    Next i
    Combo1.Top = alto
    btnBuscar(2).Top = alto
    ' ***************************************************
End Sub

Private Sub BotonEliminar()
Dim sql As String
Dim temp As Boolean

    On Error GoTo Error2

    'Certes comprovacions
    If adodc1.Recordset.EOF Then Exit Sub
'    If Not SepuedeBorrar Then Exit Sub

    ' *** repasar el nom de l'adodc, l'index del Field i el camp que te la PK ***
    ' El registre de codi 0 no es pot Modificar ni Eliminar
    ' If EsCodigoCero(CStr(adodc1.Recordset.Fields(1).Value), FormatoCampo(txtAux(1))) Then Exit Sub
    ' ***************************************************************************

    '*** canviar la pregunta, els noms dels camps i el DELETE; repasar codEmpre ***
    sql = "¿Seguro que desea eliminar el Dato del Turno?"
    'SQL = SQL & vbCrLf & "Código: " & Format(adodc1.Recordset.Fields(0), "000")
    sql = sql & vbCrLf & "Fecha: " & adodc1.Recordset.Fields(0)
    sql = sql & vbCrLf & "Turno: " & adodc1.Recordset.Fields(1)
    sql = sql & vbCrLf & "Tipo: " & adodc1.Recordset.Fields(4)
    
    If MsgBox(sql, vbQuestion + vbYesNo) = vbYes Then
        'N'hi ha que eliminar
        NumRegElim = adodc1.Recordset.AbsolutePosition
        sql = "Delete from sturno where fechatur = " & DBSet(adodc1.Recordset!fechatur, "F") & " AND codturno = " & adodc1.Recordset!codTurno & " AND numlinea = " & adodc1.Recordset!NumLinea
        Conn.Execute sql
    '******************************************************************************
        CargaGrid CadB
        temp = SituarDataTrasEliminar(adodc1, NumRegElim, True)
        PonerModoOpcionesMenu
        adodc1.Recordset.Cancel
    End If
    Exit Sub
    
Error2:
    Screen.MousePointer = vbDefault
    MuestraError Err.Number, "Eliminando registro", Err.Description
End Sub

Private Sub PonerLongCampos()
    PonerLongCamposGnral Me, Modo, 3
End Sub

Private Sub cmdAceptar_Click()
Dim i As Long

    Select Case Modo
        Case 3 'INSERTAR
            If DatosOk Then
                txtAux(2) = SugerirCodigoSiguienteStr("sturno", "numlinea", "fechatur=" & DBSet(txtAux(0).Text, "F") & " AND codturno=" & DBSet(txtAux(1).Text, "N"))
                If InsertarDesdeForm2(Me, 0) Then
                    CargaGrid
                    If (DatosADevolverBusqueda <> "") And NuevoCodigo <> "" Then
                        cmdCancelar_Click
                        If Not adodc1.Recordset.EOF Then
                            ' *** filtrar per tota la PK; repasar codEmpre **
                            adodc1.Recordset.Filter = "fechatur = " & txtAux(0).Text & " AND codturno = " & txtAux(1).Text
                            ' ****************************************************
                        End If
                        cmdRegresar_Click
                    Else
                        BotonAnyadir
                    End If
                    CadB = ""
                End If
            End If

        Case 4 'MODIFICAR
            If DatosOk Then
                'If ModificaDesdeFormulario(Me) Then
                If ModificaDesdeFormulario2(Me, 0) Then
                    i = adodc1.Recordset.AbsolutePosition
                    TerminaBloquear
                    PonerModo 2
                    CargaGrid CadB
                    adodc1.Recordset.Move i - 1
                    PonerFocoGrid Me.DataGrid1
                End If
            End If

        Case 1  'BUSQUEDA
            CadB = ObtenerBusqueda(Me)
            If CadB <> "" Then
                CargaGrid CadB
                PonerModo 2
                PonerFocoGrid Me.DataGrid1
            End If
    End Select
End Sub

Private Sub cmdCancelar_Click()
'On Error Resume Next

    Select Case Modo
        Case 3 'INSERTAR
            DataGrid1.AllowAddNew = False
            'CargaGrid
            If Not adodc1.Recordset.EOF Then adodc1.Recordset.MoveFirst
        Case 4 'MODIFICAR
            TerminaBloquear
        Case 1 'BUSQUEDA
            CargaGrid CadB
    End Select

    If Not adodc1.Recordset.EOF Then
        CargaForaGrid
    Else
        LimpiarCampos
    End If

    PonerModo 2
    PonerFocoGrid Me.DataGrid1

End Sub

Private Sub cmdRegresar_Click()
Dim cad As String
Dim i As Integer
Dim J As Integer
Dim Aux As String

    If adodc1.Recordset.EOF Then
        MsgBox "Ningún registro devuelto.", vbExclamation
        Exit Sub
    End If
    cad = ""
    i = 0
    Do
        J = i + 1
        i = InStr(J, DatosADevolverBusqueda, "|")
        If i > 0 Then
            Aux = Mid(DatosADevolverBusqueda, J, i - J)
            J = Val(Aux)
            cad = cad & adodc1.Recordset.Fields(J) & "|"
        End If
    Loop Until i = 0
    RaiseEvent DatoSeleccionado(cad)
    Unload Me
End Sub

Private Sub DataGrid1_DblClick()
    If cmdRegresar.visible Then cmdRegresar_Click
End Sub

Private Sub DataGrid1_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub Form_Activate()
    Screen.MousePointer = vbDefault
    'Posem el foco
    If (DatosADevolverBusqueda <> "") And NuevoCodigo <> "" Then
        PonerFoco txtAux(0)
    End If
End Sub

Private Sub Form_Load()
Dim i As Integer

    '******* repasar si n'hi ha botó d'imprimir o no******
    With Me.Toolbar1
        .HotImageList = frmPpal.imgListComun_OM
        .DisabledImageList = frmPpal.imgListComun_BN
        .ImageList = frmPpal.imgListComun
        'el 1 es separadors
        .Buttons(2).Image = 1   'Buscar
        .Buttons(3).Image = 2   'Tots
        'el 4 i el 5 son separadors
        .Buttons(6).Image = 3   'Insertar
        .Buttons(7).Image = 4   'Modificar
        .Buttons(8).Image = 5   'Borrar
        .Buttons(9).Image = 16   'Generacion masiva
        
        'el 10 son separadors
        .Buttons(11).Image = 10  'Imprimir
        .Buttons(12).Image = 11  'Eixir
    End With
    '*****************************************************

    'IMAGES para busqueda
    For i = 0 To Me.imgBuscar.Count - 1
        Me.imgBuscar(i).Picture = frmPpal.imgListImages16.ListImages(1).Picture
    Next i

    chkVistaPrevia.Value = CheckValueLeer(Name)
    ' *** SI N'HI HAN COMBOS ***
    CargaCombo
    ' **************************

    '****************** canviar la consulta *********************************+
    CadenaConsulta = "SELECT sturno.fechatur, sturno.codturno, sturno.numlinea, "
    CadenaConsulta = CadenaConsulta & "sturno.tiporegi, tiporegi.nomtipre, "
    CadenaConsulta = CadenaConsulta & "sturno.numtanqu, sturno.nummangu, sturno.codartic, "
    CadenaConsulta = CadenaConsulta & "sartic.nomartic, sturno.tipocred, tipocred.nomcredi, "
    CadenaConsulta = CadenaConsulta & "sturno.litrosve, sturno.importel, sturno.containi, "
    CadenaConsulta = CadenaConsulta & "sturno.contafin "
    CadenaConsulta = CadenaConsulta & " FROM sturno, tiporegi, sartic, tipocred "
    CadenaConsulta = CadenaConsulta & " WHERE sturno.tiporegi = tiporegi.tiporegi "
    CadenaConsulta = CadenaConsulta & " AND   sturno.codartic = sartic.codartic "
    CadenaConsulta = CadenaConsulta & " AND   sturno.tipocred = tipocred.tipocred "
    '************************************************************************

    CadB = ""
    CargaGrid

    ' ****** Si n'hi han camps fora del grid ******
    ' *** NOTA: açò, no se per què, ara no fa falta ***
    'CargaForaGrid
    ' *********************************************

    If (DatosADevolverBusqueda <> "") And NuevoCodigo <> "" Then
        BotonAnyadir
    Else
        PonerModo 2
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    CheckValueGuardar Me.Name, Me.chkVistaPrevia.Value
    Screen.MousePointer = vbDefault
End Sub

Private Sub frmTCred_DatoSeleccionado(CadenaSeleccion As String)
    txtAux(6).Text = RecuperaValor(CadenaSeleccion, 1)
    FormateaCampo txtAux(6)
    txtAux2(6).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub mnBuscar_Click()
    BotonBuscar
End Sub

Private Sub mnEliminar_Click()
    BotonEliminar
End Sub

Private Sub mnGeneracion_Click()
    frmListado.OpcionListado = 14
    frmListado.Show vbModal
    CargaGrid
End Sub

Private Sub mnModificar_Click()
    'Comprobaciones
    '--------------
    If adodc1.Recordset.EOF Then Exit Sub
    If adodc1.Recordset.RecordCount < 1 Then Exit Sub
    
    ' *** repasar el nom de l'adodc, l'index del Field i el camp que te la PK ***
    'El registre de codi 0 no es pot Modificar ni Eliminar
    'If EsCodigoCero(CStr(adodc1.Recordset.Fields(1).Value), FormatoCampo(txtAux(1))) Then Exit Sub
    ' ***************************************************************************

    
    'Prepara para modificar
    '----------------------
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

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 2: mnBuscar_Click
        Case 3: mnVerTodos_Click
    
        Case 6: mnNuevo_Click
        Case 7: mnModificar_Click
        Case 8: mnEliminar_Click

        Case 9: mnGeneracion_Click
        
'        Case 11: printNou
        Case 12: mnSalir_Click
    End Select
End Sub

Private Sub CargaGrid(Optional vSQL As String)
    Dim i As Integer
    Dim sql As String
    Dim tots As String

'    adodc1.ConnectionString = Conn
    ' *** si en Form_load ya li he posat clausula WHERE, canviar el `WHERE` de
    ' `SQL = CadenaConsulta & " WHERE " & vSQL` per un `AND`
    If vSQL <> "" Then
        sql = CadenaConsulta & " AND " & vSQL
    Else
        sql = CadenaConsulta
    End If
    '********************* canviar el ORDER BY *********************++
    sql = sql & " ORDER BY fechatur desc, codturno desc, numlinea "

    CargaGridGnral Me.DataGrid1, Me.adodc1, sql, False

    ' *** posar només els controls del grid ***
    tots = "S|txtAux(0)|T|Fecha|1200|;S|btnBuscar(2)|B||195|;S|txtAux(1)|T|Turno|650|;N|||||;"
    tots = tots & "N|||||;S|Combo1|C|T.Registro|1300|;"
    tots = tots & "S|txtAux(3)|T|Tanque|900|;S|txtAux(4)|T|Manguera|950|;"
    For i = 1 To 8
        tots = tots & "N||||0|;"
    Next i
    arregla tots, DataGrid1, Me
    DataGrid1.ScrollBars = dbgAutomatic
    ' **********************************************************

    ' *** alliniar les columnes que siguen numèriques a la dreta ***
    DataGrid1.Columns(1).Alignment = dbgRight
    DataGrid1.Columns(5).Alignment = dbgRight
    DataGrid1.Columns(6).Alignment = dbgRight
    ' *****************************

    ' *** Si n'hi han camps fora del grid ***
    If Not adodc1.Recordset.EOF Then
        CargaForaGrid
    Else
        LimpiarCampos
    End If
    ' **************************************
End Sub

Private Sub txtAux_GotFocus(Index As Integer)
'   If (Modo <> 0 And Modo <> 2) Then (PARA NO VER AZULITO)
    ConseguirFocoLin txtAux(Index)
End Sub

Private Sub txtAux_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    KEYdown KeyCode
End Sub

Private Sub txtAux_KeyPress(Index As Integer, KeyAscii As Integer)
'14/02/2007 lo he comentado y lo he cambiado por lo de abajo
'    If Index = 3 And KeyAscii = 43 Then '+
'        KeyAscii = 0
''        btnBuscar_Click (1)
'    Else
'        KEYpress KeyAscii
'    End If
    If KeyAscii = teclaBuscar Then
        If Modo = 1 Or Modo = 3 Or Modo = 4 Then
            KeyAscii = 0
            Select Case Index
                Case 0: btnBuscar_Click (2) ' fecha
                Case 5: imgBuscar_Click (0) ' articulo
                Case 6: imgBuscar_Click (1) ' tipo de credito
            End Select
        End If
    Else
        KEYpress KeyAscii
    End If



End Sub

Private Sub txtAux_LostFocus(Index As Integer)
Dim cadMen As String

    If Not PerderFocoGnral(txtAux(Index), Modo) Then Exit Sub
    
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub

    '*** configurar el LostFocus dels camps (de dins i de fora del grid) ***
    Select Case Index
        Case 0 'fecha
            If txtAux(Index).Text <> "" Then PonerFormatoFecha txtAux(Index)
        
        Case 2, 3, 4
            PonerFormatoEntero txtAux(Index)
            
        Case 7, 8, 9, 10 'Importe
            If txtAux(Index).Text <> "" Then PonerFormatoDecimal txtAux(Index), 1 'tipo1: Decimal(12,2)
        
        Case 5 'Articulo
            If PonerFormatoEntero(txtAux(Index)) Then
                txtAux2(Index).Text = PonerNombreDeCod(txtAux(Index), "sartic", "nomartic")
                If txtAux2(Index).Text = "" Then
                    cadMen = "No existe el Articulo: " & txtAux(Index).Text & vbCrLf
                    cadMen = cadMen & "¿Desea crearlo?" & vbCrLf
                    If MsgBox(cadMen, vbQuestion + vbYesNo) = vbYes Then
                        Set frmArt = New frmManArtic
                        frmArt.DatosADevolverBusqueda = "0|1|"
                        frmArt.NuevoCodigo = txtAux(Index).Text
                        txtAux(Index).Text = ""
                        TerminaBloquear
                        frmArt.Show vbModal
                        Set frmArt = Nothing
                        If Modo = 4 Then BLOQUEADesdeFormulario2 Me, adodc1, 1
                    Else
                        txtAux(Index).Text = ""
                    End If
                    PonerFoco txtAux(Index)
                End If
            Else
                txtAux2(Index).Text = ""
            End If
            
        Case 6 'TIPO CREDITO
            If PonerFormatoEntero(txtAux(Index)) Then
                txtAux2(Index).Text = PonerNombreDeCod(txtAux(Index), "tipocred", "nomcredi")
                If txtAux2(Index).Text = "" Then
                    cadMen = "No existe el Tipo Credito: " & txtAux(Index).Text & vbCrLf
                    MsgBox cadMen, vbInformation
                    txtAux(Index).Text = ""
                    PonerFoco txtAux(Index)
                End If
            Else
                txtAux2(Index).Text = ""
            End If
            
    End Select
End Sub

Private Function DatosOk() As Boolean
Dim Datos As String
Dim b As Boolean
' *** només per ad este manteniment ***
Dim RS As Recordset
Dim cad As String
' *************************************

    b = CompForm(Me)
    If Not b Then Exit Function

     If b And (Modo = 3) Then
        'Estem insertant
        'aço es com posar: select codvarie from svarie where codvarie = txtAux(0)
        'la N es pa dir que es numèric

        ' *** canviar els arguments de la funcio, el mensage i repasar si n'hi ha codEmpre ***
'        Datos = DevuelveDesdeBDNew(1, "sturno", "fechatur", "fechatur", txtAux(0).Text, "F", "", "codturno", txtAux(1).Text, "N", "numlinea", txtAux(2).Text, "N")
'
'        If Datos <> "" Then
'            MsgBox "Ya existe el Turno de esa Fecha: " & txtAux(0).Text, vbExclamation
'            DatosOk = False
'            PonerFoco txtAux(0)
'            Exit Function
'        End If
        '*************************************************************************************
     End If

    
    DatosOk = b
End Function

Private Sub PonerOpcionesMenu()
    PonerOpcionesMenuGeneral Me
    PonerOpcionesMenuGeneralNew Me
End Sub

Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, Modo, cerrar
    If cerrar Then Unload Me
End Sub

' ********** SI N'HI HAN CAMPS FORA DEL GRID ******************

Private Sub DataGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
Dim i As Byte

    If Modo <> 4 Then 'Modificar
        CargaForaGrid
'        Me.lblIndicador = PonerContRegistros(Me.adodc1)
    Else
        For i = 0 To txtAux.Count - 1
            txtAux(i).Text = ""
        Next i
    End If
    
    PonerContRegIndicador
    
End Sub

Private Sub CargaForaGrid()

    If DataGrid1.Columns.Count > 2 Then
        ' *** posar als camps de fora del grid el valor de la columna corresponent ***
        txtAux(5).Text = DataGrid1.Columns(7).Text
        txtAux2(5).Text = DataGrid1.Columns(8).Text
        txtAux(6).Text = DataGrid1.Columns(9).Text
        txtAux2(6).Text = DataGrid1.Columns(10).Text
        txtAux(7).Text = DataGrid1.Columns(11).Text
        txtAux(8).Text = DataGrid1.Columns(12).Text
        txtAux(9).Text = DataGrid1.Columns(13).Text
        txtAux(10).Text = DataGrid1.Columns(14).Text

        ' *** formateje tots els camps de fora del grid ***
'        If txtAux(0).Text <> "" Then PonerFormatoFecha txtAux(0)
 '       PonerFormatoEntero txtAux(1)
  '      PonerFormatoEntero txtAux(3)
        PonerFormatoEntero txtAux(5)
        If txtAux(7).Text <> "" Then PonerFormatoDecimal txtAux(7), 3 'tipo3: Decimal(10,2)
        If txtAux(8).Text <> "" Then PonerFormatoDecimal txtAux(8), 3 'tipo3: Decimal(10,2)
        If txtAux(9).Text <> "" Then PonerFormatoDecimal txtAux(9), 3 'tipo3: Decimal(10,2)
        If txtAux(10).Text <> "" Then PonerFormatoDecimal txtAux(10), 3 'tipo3: Decimal(10,2)
        ' *************************************************
    End If
End Sub

Private Sub LimpiarCampos()
Dim i As Integer

    On Error Resume Next

    ' *** posar a huit tots els camps de fora del grid ***
    For i = 5 To 10
        txtAux(i).Text = ""
    Next i
    txtAux2(5).Text = ""
    txtAux2(6).Text = ""
    ' ****************************************************

    If Err.Number <> 0 Then Err.Clear
End Sub
' ******************************************************************

Private Sub frmArt_DatoSeleccionado(CadenaSeleccion As String)
    txtAux(5).Text = RecuperaValor(CadenaSeleccion, 1)
    FormateaCampo txtAux(5)
    txtAux2(5).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmC_Selec(vFecha As Date)
    ' *** repasar si el camp es txtAux o Text1 ***
    txtAux(0).Text = Format(vFecha, "dd/mm/yyyy") '<===
    ' ********************************************
End Sub

' *** si n'hi han botons de buscar codi fora del grid ***
Private Sub imgBuscar_Click(Index As Integer)
    TerminaBloquear
    
    Select Case Index
        Case 0 'articulo
            Set frmArt = New frmManArtic
            frmArt.DatosADevolverBusqueda = "0|1|"
            frmArt.Show vbModal
            Set frmArt = Nothing
            PonerFoco txtAux(5)
        Case 1 ' tipo de credito
            Set frmTCred = New frmManTipCre
            frmTCred.DatosADevolverBusqueda = "0|1|"
            frmTCred.DeConsulta = True
            frmTCred.Show vbModal
            Set frmTCred = Nothing
            PonerFoco txtAux(6)
    
    End Select
    If Modo = 4 Then BLOQUEADesdeFormulario2 Me, adodc1, 1
    
 End Sub
 
 Private Sub btnBuscar_Click(Index As Integer)
    TerminaBloquear
    Select Case Index
        Case 2 ' Fecha
            Dim esq As Long
            Dim dalt As Long
            Dim menu As Long
            Dim obj As Object
        
            Set frmC = New frmCal
            
            esq = btnBuscar(Index).Left
            dalt = btnBuscar(Index).Top
                
            Set obj = btnBuscar(Index).Container
              
              While btnBuscar(Index).Parent.Name <> obj.Name
                    esq = esq + obj.Left
                    dalt = dalt + obj.Top
                    Set obj = obj.Container
              Wend
            
            menu = Me.Height - Me.ScaleHeight 'ací tinc el heigth del menú i de la toolbar
        
            frmC.Left = esq + btnBuscar(Index).Parent.Left + 30
            frmC.Top = dalt + btnBuscar(Index).Parent.Top + btnBuscar(Index).Height + menu - 40
        
            btnBuscar(Index).Tag = Index '<===
            ' *** repasar si el camp es txtAux o Text1 ***
            If txtAux(0).Text <> "" Then frmC.NovaData = txtAux(0).Text
            ' ********************************************
        
            frmC.Show vbModal
            Set frmC = Nothing
            ' *** repasar si el camp es txtAux o Text1 ***
            PonerFoco txtAux(0) '<===
            ' ********************************************
        
            
'        Case 1 'lugar de recogida
'            ' *** adrede per ad este manteniment ***
'            If (txtAux(4).Text = "" Or Not IsNumeric(txtAux(4).Text)) Then
'                MsgBox "Debe seleccionar una población correcta.", vbExclamation
'                Exit Sub
'            End If
'            ' *************************************
'            Set frmPobr = New frmPobrecog
'            frmPobr.DatosADevolverBusqueda = "1|2|" 'no pose el 0 per a no retornar el codpobla
'            frmPobr.Poblacio = txtAux(4).Text '*** adrede per ad este manteniment ***
'            frmPobr.Show vbModal
'            Set frmPobr = Nothing
'            PonerFoco txtAux(5)
            
    End Select
    
    If Modo = 4 Then BLOQUEADesdeFormulario2 Me, adodc1, 1
End Sub

Private Sub PonerContRegIndicador()
'si estamos en modo ver registros muestra el numero de registro en el que estamos
'situados del total de registros mostrados: 1 de 24
Dim cadReg As String

    If (Modo = 2 Or Modo = 0) Then
        cadReg = PonerContRegistros(Me.adodc1)
        If CadB = "" Then
            lblIndicador.Caption = cadReg
        Else
            lblIndicador.Caption = "BUSQUEDA: " & cadReg
        End If
    End If
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub CargaCombo()
Dim cad As String
Dim RS As ADODB.Recordset

    On Error GoTo ErrCarga
    Combo1.Clear
    
    cad = "SELECT * FROM tiporegi ORDER BY tiporegi"
    Set RS = New ADODB.Recordset
    RS.Open cad, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    
    'Combo1.AddItem "" 'pose uno en blanc sinse valor
    While Not RS.EOF
        Combo1.AddItem RS!nomtipre
        Combo1.ItemData(Combo1.NewIndex) = RS!tiporegi
        RS.MoveNext
        '.MoveNext
    Wend
    RS.Close
    Set RS = Nothing
    
    Exit Sub
    
ErrCarga:
    If Err.Number <> 0 Then MuestraError Err.Number, "Cargar datos combo.", Err.Description
End Sub

' ### [DavidV] 26/04/2006: Activar/desactivar la rueda del ratón.
Private Sub DataGrid1_GotFocus()
'  WheelHook DataGrid1
End Sub
Private Sub DataGrid1_Lostfocus()
'  WheelUnHook
End Sub
