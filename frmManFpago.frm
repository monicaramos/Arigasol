VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmManFpago 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Formas de Pago"
   ClientHeight    =   5925
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13485
   Icon            =   "frmManFpago.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5925
   ScaleWidth      =   13485
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox Combo1 
      Appearance      =   0  'Flat
      Height          =   315
      Index           =   4
      ItemData        =   "frmManFpago.frx":000C
      Left            =   11100
      List            =   "frmManFpago.frx":000E
      Style           =   2  'Dropdown List
      TabIndex        =   12
      Tag             =   " FP Vale|N|N|||sforpa|tipovale|0||"
      Top             =   2280
      Width           =   2175
   End
   Begin VB.TextBox txtAux2 
      BackColor       =   &H80000018&
      Enabled         =   0   'False
      Height          =   285
      Index           =   8
      Left            =   10170
      TabIndex        =   22
      Top             =   1830
      Width           =   3105
   End
   Begin VB.TextBox txtAux 
      Height          =   290
      Index           =   8
      Left            =   11100
      MaxLength       =   10
      TabIndex        =   11
      Tag             =   "Codigo Cliente|N|S|||sforpa|codsocio|000000||"
      Top             =   1440
      Width           =   735
   End
   Begin VB.TextBox txtAux 
      Height          =   290
      Index           =   7
      Left            =   11520
      MaxLength       =   10
      TabIndex        =   10
      Tag             =   "Codigo Ext|T|S|||sforpa|codexterno|||"
      Top             =   1020
      Width           =   735
   End
   Begin VB.TextBox txtAux 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   290
      Index           =   6
      Left            =   6750
      MaxLength       =   3
      TabIndex        =   8
      Tag             =   "Resto Vto|N|N|0|99|sforpa|restoven|##0||"
      Top             =   4980
      Width           =   405
   End
   Begin VB.TextBox txtAux 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   290
      Index           =   5
      Left            =   5670
      MaxLength       =   2
      TabIndex        =   6
      Tag             =   "Nro Vtos|N|N|0|99|sforpa|numerove|#0||"
      Top             =   4950
      Width           =   525
   End
   Begin VB.TextBox txtAux 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   290
      Index           =   4
      Left            =   7230
      MaxLength       =   2
      TabIndex        =   9
      Tag             =   "Forpa Alvic|N|N|0|99|sforpa|forpaalvic|00||"
      Top             =   4980
      Width           =   735
   End
   Begin VB.TextBox txtAux 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   290
      Index           =   3
      Left            =   6240
      MaxLength       =   2
      TabIndex        =   7
      Tag             =   "Dias Vto|N|N|0|9999|sforpa|diasvto|##0||"
      Top             =   4950
      Width           =   465
   End
   Begin VB.ComboBox Combo1 
      Appearance      =   0  'Flat
      Height          =   315
      Index           =   3
      ItemData        =   "frmManFpago.frx":0010
      Left            =   4920
      List            =   "frmManFpago.frx":0012
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Tag             =   "Permite bonificacion|N|N|0|1|sforpa|permitebonif|||"
      Top             =   4920
      Width           =   735
   End
   Begin VB.ComboBox Combo1 
      Appearance      =   0  'Flat
      Height          =   315
      Index           =   2
      ItemData        =   "frmManFpago.frx":0014
      Left            =   4080
      List            =   "frmManFpago.frx":0016
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Tag             =   "Se contabiliza|N|N|0|1|sforpa|contabilizasn|||"
      Top             =   4920
      Width           =   735
   End
   Begin VB.TextBox txtAux2 
      BackColor       =   &H80000018&
      Enabled         =   0   'False
      Height          =   285
      Index           =   2
      Left            =   10200
      TabIndex        =   21
      Top             =   3090
      Width           =   3135
   End
   Begin VB.TextBox txtAux 
      Height          =   290
      Index           =   2
      Left            =   11910
      MaxLength       =   10
      TabIndex        =   13
      Tag             =   "Cuenta Contable|T|S|||sforpa|codmacta|||"
      Top             =   2760
      Width           =   1395
   End
   Begin VB.ComboBox Combo1 
      Appearance      =   0  'Flat
      Height          =   315
      Index           =   1
      ItemData        =   "frmManFpago.frx":0018
      Left            =   3270
      List            =   "frmManFpago.frx":001A
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Tag             =   "Sale cuadre|N|N|0|1|sforpa|cuadresn|||"
      Top             =   4920
      Width           =   735
   End
   Begin VB.ComboBox Combo1 
      Appearance      =   0  'Flat
      Height          =   315
      Index           =   0
      ItemData        =   "frmManFpago.frx":001C
      Left            =   2400
      List            =   "frmManFpago.frx":001E
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Tag             =   "Tipo F.Pago|N|N|0|5|sforpa|tipforpa|||"
      Top             =   4920
      Width           =   735
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   6870
      TabIndex        =   14
      Top             =   5490
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   8130
      TabIndex        =   15
      Top             =   5460
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   290
      Index           =   1
      Left            =   900
      MaxLength       =   30
      TabIndex        =   1
      Tag             =   "Descripción de F.Pago|T|N|||sforpa|nomforpa|||"
      Top             =   4920
      Width           =   1395
   End
   Begin VB.TextBox txtAux 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   290
      Index           =   0
      Left            =   60
      MaxLength       =   3
      TabIndex        =   0
      Tag             =   "Código de F.Pago|N|N|0|99|sforpa|codforpa|00|S|"
      Top             =   4920
      Width           =   800
   End
   Begin VB.CommandButton cmdRegresar 
      Caption         =   "&Regresar"
      Height          =   375
      Left            =   8130
      TabIndex        =   20
      Top             =   5460
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Height          =   555
      Index           =   1
      Left            =   120
      TabIndex        =   16
      Top             =   5310
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
         TabIndex        =   17
         Top             =   240
         Width           =   2295
      End
   End
   Begin MSAdodcLib.Adodc adodc1 
      Height          =   375
      Left            =   3960
      Top             =   0
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   661
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
      TabIndex        =   18
      Top             =   0
      Width           =   13485
      _ExtentX        =   23786
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
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
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
         TabIndex        =   19
         Top             =   120
         Visible         =   0   'False
         Width           =   1215
      End
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmManFpago.frx":0020
      Height          =   4410
      Left            =   90
      TabIndex        =   28
      Top             =   630
      Width           =   10035
      _ExtentX        =   17701
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
   Begin VB.Image imgBuscar 
      Height          =   270
      Index           =   0
      Left            =   11310
      ToolTipText     =   "Buscar Colectivo"
      Top             =   2790
      Width           =   270
   End
   Begin VB.Image imgBuscar 
      Height          =   270
      Index           =   1
      Left            =   10770
      ToolTipText     =   "Buscar Cliente"
      Top             =   1440
      Width           =   270
   End
   Begin VB.Label Label4 
      Caption         =   "Cta.Contable"
      Height          =   285
      Index           =   4
      Left            =   10230
      TabIndex        =   27
      ToolTipText     =   "Buscar Cuenta"
      Top             =   2760
      Width           =   1035
   End
   Begin VB.Label Label4 
      Caption         =   "Tipo  Pago"
      Height          =   285
      Index           =   3
      Left            =   10170
      TabIndex        =   26
      Top             =   2310
      Width           =   1245
   End
   Begin VB.Label Label4 
      Caption         =   "Cliente"
      Height          =   255
      Index           =   2
      Left            =   10200
      TabIndex        =   25
      Top             =   1470
      Width           =   525
   End
   Begin VB.Label Label4 
      Caption         =   "Código Externo"
      Height          =   255
      Index           =   1
      Left            =   0
      TabIndex        =   24
      Top             =   0
      Width           =   1245
   End
   Begin VB.Label Label4 
      Caption         =   "Código Externo"
      Height          =   255
      Index           =   0
      Left            =   10170
      TabIndex        =   23
      Top             =   1050
      Width           =   1245
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
      Begin VB.Menu mnBarra2 
         Caption         =   "-"
      End
      Begin VB.Menu mnImprimir 
         Caption         =   "&Imprimir"
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
Attribute VB_Name = "frmManFpago"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' +-+-+-+-+-+-+-+-+-+-+-+-
' +-+- Autor: MANOLO  +-+-
' +-+-+-+-+-+-+-+-+-+-+-+-

' **************** PER A QUE FUNCIONE EN UN ATRE MANTENIMENT ********************
' 0. Posar-li l'atribut Datasource a "adodc1" del Datagrid1. Canviar el Caption
'    del formulari
' 1. Canviar els TAGs i els Maxlength de TextAux(0) i TextAux(1)
' 2. En PonerModo(vModo) repasar els indexs del botons, per si es canvien
' 3. En la funció BotonAnyadir() canviar la taula i el camp per a SugerirCodigoSiguienteStr
' 4. En la funció BotonBuscar() canviar el nom de la clau primaria
' 5. En la funció BotonEliminar() canviar la pregunta, les descripcions de la
'    variable SQL i el contingut del DELETE
' 6. En la funció PonerLongCampos() posar els camps als que volem canviar el MaxLength quan busquem
' 7. En Form_Load() repasar la barra d'iconos (per si es vol canviar algún) i
'    canviar la consulta per a vore tots els registres
' 8. En Toolbar1_ButtonClick repasar els indexs de cada botó per a que corresponguen
' 9. En la funció CargaGrid canviar l'ORDER BY (normalment per la clau primaria);
'    canviar ademés els noms dels camps, el format i si fa falta la cantitat;
'    repasar els index dels botons modificar i eliminar.
'    NOTA: si en Form_load ya li he posat clausula WHERE, canviar el `WHERE` de
'    `SQL = CadenaConsulta & " WHERE " & vSQL` per un `AND`
' 10. En txtAux_LostFocus canviar el mensage i el format del camp
' 11. En la funció DatosOk() canviar els arguments de DevuelveDesdeBD i el mensage
'    en cas d'error
' 12. En la funció SepuedeBorrar() canviar les comprovacions per a vore si es pot
'    borrar el registre
' *******************************SI N'HI HA COMBO*******************************
' 0. Comprovar que en el SQL de Form_Load() es faça referència a la taula del Combo
' 1. Pegar el Combo1 al  costat dels TextAux. Canviar-li el TAG
' 2. En BotonModificar() canviar el camp del Combo
' 3. En CargaCombo() canviar la consulta i els noms del camps, o posar els valor
'    a ma si no es llig de cap base de datos els valors del Combo

Option Explicit

Public DatosADevolverBusqueda As String    'Tendra el nº de text que quiere que devuelva, empipados
Public Event DatoSeleccionado(CadenaSeleccion As String)

'codi per al registe que s'afegix al cridar des d'atre formulari.
'Obrir en modo Insertar i tornar datos del registre insertat
Public NuevoCodigo As String
Public CodigoActual As String
Public DeConsulta As Boolean

Private CadenaConsulta As String
Private CadB As String

Private WithEvents frmCtas As frmCtasConta 'cuentas contables
Attribute frmCtas.VB_VarHelpID = -1
Private WithEvents frmB As frmBuscaGrid
Attribute frmB.VB_VarHelpID = -1
Dim Modo As Byte
'----------- MODOS --------------------------------
'   0.-  Formulari net sense cap camp ple
'   1.-  Preparant per a fer la búsqueda
'   2.-  Ja tenim registres i els anem a recorrer
'        ,podem editar-los Edició del camp
'   3.-  Inserció de nou registre
'   4.-  Modificar
'--------------------------------------------------
Dim PrimeraVez As Boolean
Dim indice As Byte 'Index del text1 on es poses els datos retornats des d'atres Formularis de Mtos
Dim I As Integer


Dim TipFPAnt As String
Dim NroVtoAnt As String
Dim DiasVtoAnt As String
Dim RestoVtoAnt As String


Private Sub PonerModo(vModo)
Dim b As Boolean

    Modo = vModo
    
    b = (Modo = 2)
    If b Then
        PonerContRegIndicador
    Else
        PonerIndicador lblIndicador, Modo
    End If
    
    For I = 0 To 3
        If I <> 2 Then
            txtAux(I).visible = Not b
        End If
        Combo1(I).visible = Not b
    Next I
    
    txtAux(4).visible = Not b
    txtAux(5).visible = Not b
    txtAux(6).visible = Not b
    
    cmdAceptar.visible = Not b
    cmdCancelar.visible = Not b
    DataGrid1.Enabled = b
    
    For I = 0 To imgBuscar.Count - 1
        imgBuscar(I).Enabled = (Modo = 1 Or Modo = 3 Or Modo = 4)
    Next I
    
    'Si es regresar
    If DatosADevolverBusqueda <> "" Then cmdRegresar.visible = b
    
    PonerLongCampos
    PonerModoOpcionesMenu 'Activar/Desact botones de menu segun Modo
    PonerOpcionesMenu  'En funcion del usuario
    
    'Si estamos modo Modificar bloquear clave primaria
    BloquearTxt txtAux(0), (Modo = 4)
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
    Me.mnImprimir.Enabled = b
    
End Sub

Private Sub BotonAnyadir()
    Dim NumF As String
    Dim anc As Single
    
'    CargaGrid 'primer de tot carregue tot el grid
    CadB = ""
    '******************** canviar taula i camp **************************
    If (DatosADevolverBusqueda <> "") And NuevoCodigo <> "" Then
        NumF = NuevoCodigo
    Else
        NumF = SugerirCodigoSiguienteStr("sforpa", "codforpa")
    End If
    '********************************************************************
    'Situamos el grid al final
    AnyadirLinea DataGrid1, adodc1
         
    anc = DataGrid1.Top
    If DataGrid1.Row < 0 Then
        anc = anc + 206
    Else
        anc = anc + DataGrid1.RowTop(DataGrid1.Row) + 5
    End If
    txtAux(0).Text = NumF
    FormateaCampo txtAux(0)
    txtAux(1).Text = ""
    txtAux(2).Text = ""
    txtAux(4).Text = ""
    txtAux(6).Text = ""
    txtAux(7).Text = ""
    txtAux(8).Text = ""
    
    For I = 0 To Combo1.Count - 1
        Combo1(I).ListIndex = -1
    Next I
    Combo1(4).ListIndex = 0
    LLamaLineas anc, 3 'Pone el form en Modo=3, Insertar
       
    'Ponemos el foco
    PonerFoco txtAux(0)
End Sub

Private Sub BotonVerTodos()
    CadB = ""
    CargaGrid ""
    PonerModo 2
End Sub

Private Sub BotonBuscar()
    ' ***************** canviar per la clau primaria ********
    CargaGrid "sforpa.codforpa = -1"
    '*******************************************************************************
    'Buscar
    For I = 0 To txtAux.Count - 1
        txtAux(I).Text = ""
    Next I
'    PosicionarCombo Combo1, "724"
    For I = 0 To Combo1.Count - 1
        Combo1(I).ListIndex = -1
    Next I
    LLamaLineas DataGrid1.Top + 206, 1 'Pone el form en Modo=1, Buscar
    PonerFoco txtAux(0)
End Sub

Private Sub BotonModificar()
    Dim anc As Single
    Dim I As Integer
    
    Screen.MousePointer = vbHourglass
    
    If DataGrid1.Bookmark < DataGrid1.FirstRow Or DataGrid1.Bookmark > (DataGrid1.FirstRow + DataGrid1.VisibleRows - 1) Then
        I = DataGrid1.Bookmark - DataGrid1.FirstRow
        DataGrid1.Scroll 0, I
        DataGrid1.Refresh
    End If
    
    If DataGrid1.Row < 0 Then
        anc = 320
    Else
        anc = DataGrid1.RowTop(DataGrid1.Row) + 630 '545
    End If

    'Llamamos al form
    txtAux(0).Text = DataGrid1.Columns(0).Text
    txtAux(1).Text = DataGrid1.Columns(1).Text
    txtAux(2).Text = DataGrid1.Columns(19).Text
    
    txtAux(3).Text = DataGrid1.Columns(11).Text
    
    txtAux(5).Text = DataGrid1.Columns(10).Text
    txtAux(6).Text = DataGrid1.Columns(12).Text
    txtAux(4).Text = DataGrid1.Columns(13).Text
    txtAux(7).Text = DataGrid1.Columns(14).Text
    txtAux(8).Text = DataGrid1.Columns(15).Text
    txtAux2(8).Text = DataGrid1.Columns(16).Text
    
    
    
    ' ***** canviar-ho pel nom del camp del combo *********
'    SelComboBool DataGrid1.Columns(2).Text, Combo1(0)
    PosicionarCombo Combo1(0), DataGrid1.Columns(2).Text
'    SelComboBool DataGrid1.Columns(3).Text, Combo1(1)
    PosicionarCombo Combo1(1), DataGrid1.Columns(4).Text
    PosicionarCombo Combo1(2), DataGrid1.Columns(6).Text
    PosicionarCombo Combo1(3), DataGrid1.Columns(8).Text
    If DataGrid1.Columns(17).Text <> "" Then PosicionarCombo Combo1(4), DataGrid1.Columns(17).Text
    ' *****************************************************
    ' ### [Monica] 12/09/2006
    If vParamAplic.NumeroConta <> 0 Then
        txtAux2(2).Text = DataGrid1.Columns(20).Text
    Else
    
    End If

    'PosicionarCombo Me.Combo1(0), i
    'PosicionarCombo Me.Combo1(1), i

    LLamaLineas anc, 4 'Pone el form en Modo=4, Modificar
   
   
    '[Monica]25/04/2018: valores anteriores de cambio
    TipFPAnt = Combo1(0).ListIndex
    NroVtoAnt = txtAux(5).Text
    DiasVtoAnt = txtAux(3).Text
    RestoVtoAnt = txtAux(6).Text
   
   
   
    'Como es modificar
    PonerFoco txtAux(1)
    Screen.MousePointer = vbDefault
End Sub

Private Sub LLamaLineas(alto As Single, xModo As Byte)
    DeseleccionaGrid Me.DataGrid1
    PonerModo xModo
    
    'Fijamos el ancho
    For I = 0 To 3
        If I <> 2 Then txtAux(I).Top = alto
        Combo1(I).Top = alto - 15
    Next I
    txtAux(4).Top = alto
    txtAux(5).Top = alto
    txtAux(6).Top = alto
'    txtAux(7).Top = alto
'    txtAux(8).Top = alto
    
    ' ### [Monica] 12/09/2006
'    txtAux2(2).Top = alto
'    txtAux2(8).Top = alto
'    btnBuscar(0).Top = alto - 15
'    btnBuscar(1).Top = alto - 15
End Sub

Private Sub BotonEliminar()
Dim SQL As String
Dim temp As Boolean

    On Error GoTo Error2
    'Ciertas comprobaciones
    If adodc1.Recordset.EOF Then Exit Sub
    
    If Not PuedeModificarFPenContab Then Exit Sub
    
    
'    If Not SepuedeBorrar Then Exit Sub
        
    ' ### [Monica] 26/09/2006 dejamos modificar y eliminar el codigo 0
    ' *** repasar el nom de l'adodc, l'index del Field i el camp que te la PK ***
    'El registre de codi 0 no es pot Modificar ni Eliminar
    'If EsCodigoCero(CStr(adodc1.Recordset.Fields(0).Value), FormatoCampo(txtAux(0))) Then Exit Sub
    ' ***************************************************************************
    
    '*************** canviar els noms i el DELETE **********************************
    SQL = "¿Seguro que desea eliminar la Forma de Pago?"
    SQL = SQL & vbCrLf & "Código: " & adodc1.Recordset.Fields(0)
    SQL = SQL & vbCrLf & "Descripción: " & adodc1.Recordset.Fields(1)
    
    If MsgBox(SQL, vbQuestion + vbYesNo) = vbYes Then
        'Hay que eliminar
        NumRegElim = adodc1.Recordset.AbsolutePosition
        SQL = "Delete from sforpa where codforpa=" & adodc1.Recordset!Codforpa
        Conn.Execute SQL
        CargaGrid CadB
'        If CadB <> "" Then
'            CargaGrid CadB
'            lblIndicador.Caption = "BUSQUEDA: " & PonerContRegistros(Me.adodc1)
'        Else
'            CargaGrid ""
'            lblIndicador.Caption = ""
'        End If
        
        If vParamAplic.ContabilidadNueva Then
            SQL = "Delete from formapago where codforpa=" & adodc1.Recordset!Codforpa
            ConnConta.Execute SQL
        Else
            SQL = "Delete from sforpa where codforpa=" & adodc1.Recordset!Codforpa
            ConnConta.Execute SQL
        End If

        temp = SituarDataTrasEliminar(adodc1, NumRegElim, True)
        PonerModoOpcionesMenu
        adodc1.Recordset.Cancel
    End If
    Exit Sub
    
Error2:
    Screen.MousePointer = vbDefault
    MuestraError Err.Number, "Eliminando registro", Err.Description
End Sub

Private Function PuedeModificarFPenContab() As Boolean
Dim Cad As String
    PuedeModificarFPenContab = False
    Set miRsAux = New ADODB.Recordset

    

    NumRegElim = 0
    If vParamAplic.ContabilidadNueva Then
        Cad = "Select count(*) from cobros where codforpa=" & txtAux(0).Text
    Else
        Cad = "Select count(*) from scobro where codforpa=" & txtAux(0).Text
    End If
    
    miRsAux.Open Cad, ConnConta, adOpenForwardOnly, adLockPessimistic
    If Not miRsAux.EOF Then NumRegElim = NumRegElim + DBLet(miRsAux.Fields(0), "N")
    miRsAux.Close
    
    
    If vParamAplic.ContabilidadNueva Then
        Cad = "Select count(*) from pagos where codforpa=" & txtAux(0).Text
    Else
        Cad = "Select count(*) from spagop where codforpa=" & txtAux(0).Text
    End If
    
    
    miRsAux.Open Cad, ConnConta, adOpenForwardOnly, adLockPessimistic
    If Not miRsAux.EOF Then NumRegElim = NumRegElim + DBLet(miRsAux.Fields(0), "N")
    miRsAux.Close
    
    
    If NumRegElim > 0 Then
        If Modo = 4 Then
            If MsgBox("Existen " & NumRegElim & " vencimientos en la tesoreria con esa forma de pago. ¿Continuar con el proceso?", vbQuestion + vbYesNo) = vbNo Then Exit Function
        Else
            'NO DEJO CONTINUAR
            MsgBox "Existen " & NumRegElim & " vencimientos en la tesoreria con esa forma de pago", vbExclamation
            Exit Function
        End If
    End If
    
    'Si llega aqui puede seguir
    PuedeModificarFPenContab = True
End Function



Private Sub PonerLongCampos()
'Modificar el MaxLength del campo en funcion de si es modo de búsqueda o no
'para los campos que permitan introducir criterios más largos del tamaño del campo
'Si maxlength=6 para codprove, en modo busqueda Maxlenth=13 'el doble + 1
    
    'para los txtAux
    PonerLongCamposGnral Me, Modo, 3
End Sub

Private Sub MandaBusquedaPrevia(CadB As String)
Dim Cad As String
        'Llamamos a al form
        Cad = ""
        Cad = Cad & "Socio|codsocio|T||15·"
        Cad = Cad & "Nombre|nomsocio|T||80·"

        Screen.MousePointer = vbHourglass
        Set frmB = New frmBuscaGrid
        frmB.vCampos = Cad
        frmB.vTabla = "ssocio"
        frmB.vSQL = "" 'cad
        frmB.vDevuelve = "0|1|" 'Campos de la tabla que devuelve
        frmB.vTitulo = "Socios"
        frmB.vSelElem = 1
        frmB.Show vbModal
        Set frmB = Nothing
        
End Sub



Private Sub cmdAceptar_Click()
    Dim I As Integer

    Select Case Modo
        Case 1 'BUSQUEDA
            CadB = ObtenerBusqueda(Me)
            If CadB <> "" Then
                CargaGrid CadB
                PonerModo 2
'                lblIndicador.Caption = "BUSQUEDA: " & PonerContRegistros(Me.adodc1)
                PonerFocoGrid Me.DataGrid1
            End If
            
        Case 3 'INSERTAR
            If DatosOk Then
                'If InsertarDesdeForm(Me) Then
                If InsertarRegistro Then
                    CargaGrid
                    If (DatosADevolverBusqueda <> "") And NuevoCodigo <> "" Then
                        cmdCancelar_Click
'                        If Not adodc1.Recordset.EOF Then adodc1.Recordset.MoveLast
                        If Not adodc1.Recordset.EOF Then
                            adodc1.Recordset.Find (adodc1.Recordset.Fields(0).Name & " =" & NuevoCodigo)
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
'                If ModificaDesdeFormulario(Me) Then
                If ModificaRegistro Then
                    TerminaBloquear
                    I = adodc1.Recordset.Fields(0)
                    PonerModo 2
                    CargaGrid CadB
'                    If CadB <> "" Then
'                        CargaGrid CadB
'                        lblIndicador.Caption = "BUSQUEDA: " & PonerContRegistros(Me.adodc1)
'                    Else
'                        CargaGrid
'                        lblIndicador.Caption = ""
'                    End If
                    adodc1.Recordset.Find (adodc1.Recordset.Fields(0).Name & " =" & I)
                    PonerFocoGrid Me.DataGrid1
                End If
            End If
    End Select
End Sub
'Private Sub SelComboBool(valor As Integer, combo As ComboBox)
'    Dim i As Integer
'    Dim j As Integer
'
'    i = valor
'    For j = 0 To combo.ListCount - 1
'        If combo.ItemData(j) = i Then
'            combo.ListIndex = j
'            Exit For
'        End If
'    Next j
'End Sub

Private Sub cmdCancelar_Click()
    On Error Resume Next
    
    Select Case Modo
        Case 1 'búsqueda
            CargaGrid CadB
        Case 3 'insertar
            DataGrid1.AllowAddNew = False
            'CargaGrid
            If Not adodc1.Recordset.EOF Then adodc1.Recordset.MoveFirst
        Case 4 'modificar
            TerminaBloquear
    End Select
    
    PonerModo 2
    
'    If CadB <> "" Then
'        lblIndicador.Caption = "BUSQUEDA: " & PonerContRegistros(Me.adodc1)
''    Else
''        lblIndicador.Caption = ""
'    End If
    
    PonerFocoGrid Me.DataGrid1
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

Private Sub Combo1_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub DataGrid1_DblClick()
    If cmdRegresar.visible Then cmdRegresar_Click
End Sub

Private Sub DataGrid1_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub DataGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    PonerContRegIndicador
    CargaForaGrid
End Sub

Private Sub Form_Activate()
    Screen.MousePointer = vbDefault

    If PrimeraVez Then
        PrimeraVez = False
        If (DatosADevolverBusqueda <> "") And NuevoCodigo <> "" Then
            BotonAnyadir
        Else
            PonerModo 2
             If Me.CodigoActual <> "" Then
                SituarData Me.adodc1, "codforpa=" & CodigoActual, "", True
            End If
        End If
    End If
End Sub

Private Sub Form_Load()
    PrimeraVez = True

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
        .Buttons(11).Image = 10  'imprimir
        .Buttons(12).Image = 11  'Salir
    End With

    For I = 0 To Me.imgBuscar.Count - 1
        Me.imgBuscar(I).Picture = frmPpal.imgListImages16.ListImages(1).Picture
    Next I


    '## A mano
'    chkVistaPrevia.Value = CheckValueLeer(Name)
    
    ' ### [Monica] 12/09/2006
    ' dependiendo de si tiene o no contabilidad o no el formulario tendrá un tamaño u otro
'    If vParamAplic.NumeroConta <> 0 Then
'        Me.Width = 18100 '15600 '12985
'        Me.cmdCancelar.Left = 13410
'        Me.cmdAceptar.Left = 12150
'        Me.cmdRegresar.Left = 12090
'        Me.DataGrid1.Width = 17700 '15200
'
'    Else
'    ' no hay conexion a la contabilidad
'        Me.Width = 15300 '12800 '9895
'        Me.cmdCancelar.Left = 10440
'        Me.cmdAceptar.Left = 9180
'        Me.cmdRegresar.Left = 9120
'        Me.DataGrid1.Width = 14800 '12300
'    End If
    
    Me.Height = 6705
    
    
    
    CargaCombo
    
    '****************** canviar la consulta *********************************+
    CadenaConsulta = "SELECT sforpa.codforpa, sforpa.nomforpa, sforpa.tipforpa, "
    CadenaConsulta = CadenaConsulta & "CASE sforpa.tipforpa WHEN 0 THEN ""Efectivo"" WHEN 1 THEN ""Transferencia"" WHEN 2 THEN ""Talon"" WHEN 3 THEN ""Pagare"" "
    '[Monica]04/01/2013: Efectivos
    CadenaConsulta = CadenaConsulta & "WHEN 4 THEN ""Recibo Bancario"" WHEN 5 THEN ""Confirming"" WHEN 6 THEN ""Tarjeta"" END, "
    CadenaConsulta = CadenaConsulta & "cuadresn, CASE cuadresn WHEN 0 THEN ""No"" WHEN 1 THEN ""Si"" END, "
    CadenaConsulta = CadenaConsulta & "contabilizasn, CASE contabilizasn WHEN 0 THEN ""No"" WHEN 1 THEN ""Si"" END, "
    CadenaConsulta = CadenaConsulta & "permitebonif, CASE permitebonif WHEN 0 THEN ""No"" WHEN 1 THEN ""Si"" END, "
    CadenaConsulta = CadenaConsulta & "sforpa.numerove, sforpa.diasvto, sforpa.restoven, sforpa.forpaalvic, sforpa.codexterno, "
    CadenaConsulta = CadenaConsulta & "sforpa.codsocio, ssocio.nomsocio, "
    CadenaConsulta = CadenaConsulta & "sforpa.tipovale, CASE tipovale WHEN 0 THEN ""Normal"" WHEN 1 THEN ""Vale"" WHEN 2 THEN ""Vale"" END, "
    CadenaConsulta = CadenaConsulta & "sforpa.codmacta "
    
    ' ### [Monica] 12/09/2006
    ' en caso de haber contabilidad muestro la descripcion de la cuenta
    If vParamAplic.NumeroConta <> 0 Then
    
        If vParamAplic.ContabilidadNueva Then
            CadenaConsulta = CadenaConsulta & ", ariconta" & vParamAplic.NumeroConta & ".cuentas.nommacta "
            CadenaConsulta = CadenaConsulta & "from (sforpa left join ssocio on sforpa.codsocio = ssocio.codsocio) left join ariconta" & DBSet(vParamAplic.NumeroConta, "N")
            CadenaConsulta = CadenaConsulta & ".cuentas on (sforpa.`codmacta` = ariconta" & DBSet(vParamAplic.NumeroConta, "N") & ".cuentas.`codmacta`)"
        Else
            CadenaConsulta = CadenaConsulta & ", conta" & vParamAplic.NumeroConta & ".cuentas.nommacta "
            CadenaConsulta = CadenaConsulta & "from (sforpa left join ssocio on sforpa.codsocio = ssocio.codsocio) left join conta" & DBSet(vParamAplic.NumeroConta, "N")
            CadenaConsulta = CadenaConsulta & ".cuentas on (sforpa.`codmacta` = conta" & DBSet(vParamAplic.NumeroConta, "N") & ".cuentas.`codmacta`)"
        End If
        
    Else
    ' no hay contabilidad
        CadenaConsulta = CadenaConsulta & "FROM (sforpa left join ssocio on sforpa.codsocio = ssocio.codsocio) "
    End If
    CadenaConsulta = CadenaConsulta & " WHERE 1 = 1 "
    '************************************************************************
    
    CadB = ""
    CargaGrid
    
'    If (DatosADevolverBusqueda <> "") And NuevoCodigo <> "" Then
'        BotonAnyadir
'    Else
'        PonerModo 2
'    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
'    CheckValueGuardar Me.Name, Me.chkVistaPrevia.Value
    Screen.MousePointer = vbDefault
End Sub

Private Sub frmB_Selecionado(CadenaDevuelta As String)
'busqueda de codigo de socio
    txtAux(8).Text = RecuperaValor(CadenaDevuelta, 1)
    FormateaCampo txtAux(8)
    txtAux2(8).Text = RecuperaValor(CadenaDevuelta, 2)
End Sub


Private Sub frmCtas_DatoSeleccionado(CadenaSeleccion As String)
'Cuentas contables de la Contabilidad
    txtAux(indice).Text = RecuperaValor(CadenaSeleccion, 1) 'codmacta
    txtAux2(indice).Text = RecuperaValor(CadenaSeleccion, 2) 'des macta
End Sub


Private Sub imgBuscar_Click(Index As Integer)
 TerminaBloquear
    
    Select Case Index
        Case 0 'Cuentas Contables (de contabilidad)
            If vParamAplic.NumeroConta = 0 Then Exit Sub
            
            indice = Index + 2
            Set frmCtas = New frmCtasConta
            frmCtas.NumDigit = 0
            frmCtas.DatosADevolverBusqueda = "0|1|"
            frmCtas.CodigoActual = txtAux(indice).Text
            frmCtas.Show vbModal
            Set frmCtas = Nothing
            PonerFoco txtAux(indice)
        Case 1
'            indice = Index + 7
'            Set frmcli = New frmManClien
'            frmcli.DatosADevolverBusqueda = "0|1|"
'            frmcli.CodigoActual = txtAux(8).Text
'            frmcli.Show vbModal
'            Set frmcli = Nothing
'            PonerFoco txtAux(8)
            MandaBusquedaPrevia ""
    End Select
    
    
    If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Me.adodc1, 1

End Sub

Private Sub mnBuscar_Click()
    BotonBuscar
End Sub

Private Sub mnEliminar_Click()
    BotonEliminar
End Sub

Private Sub mnImprimir_Click()
    printNou
End Sub

Private Sub mnModificar_Click()
    'Comprobaciones
    '--------------
    If adodc1.Recordset.EOF Then Exit Sub
    
    If adodc1.Recordset.RecordCount < 1 Then Exit Sub
    
    ' ### [Monica] 26/09/2006 dejamos modificar y eliminar el codigo 0
    ' *** repasar el nom de l'adodc, l'index del Field i el camp que te la PK ***
    'El registre de codi 0 no es pot Modificar ni Eliminar
    'If EsCodigoCero(CStr(adodc1.Recordset.Fields(0).Value), FormatoCampo(txtAux(0))) Then Exit Sub
    
    
    'Preparamos para modificar
    '-------------------------
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
        Case 2
                mnBuscar_Click
        Case 3
                mnVerTodos_Click
        Case 6
                mnNuevo_Click
        Case 7
                mnModificar_Click
        Case 8
                mnEliminar_Click
        Case 11
                'MsgBox "Imprimir...under construction"
                mnImprimir_Click
        Case 12
                mnSalir_Click
    End Select
End Sub

Private Sub CargaGrid(Optional vSQL As String)
    Dim SQL As String
    Dim tots As String
    
'    adodc1.ConnectionString = Conn
    If vSQL <> "" Then
        SQL = CadenaConsulta & " AND " & vSQL
    Else
        SQL = CadenaConsulta
    End If
    '********************* canviar el ORDER BY *********************++
    SQL = SQL & " ORDER BY sforpa.codforpa"
    '**************************************************************++
    
    CargaGridGnral DataGrid1, Me.adodc1, SQL, PrimeraVez
    
    ' *******************canviar els noms i si fa falta la cantitat********************
    tots = "S|txtAux(0)|T|Cód.|500|;S|txtAux(1)|T|Descripción|2600|;"
    tots = tots & "N||||0|;S|Combo1(0)|C|Tipo|1600|;"
    tots = tots & "N||||0|;S|Combo1(1)|C|Cuadre|750|;"
    tots = tots & "N||||0|;S|Combo1(2)|C|Contab|750|;"
    tots = tots & "N||||0|;S|Combo1(3)|C|Bonif.|750|;"
    tots = tots & "S|txtAux(5)|T|Vtos|600|;"
    tots = tots & "S|txtAux(3)|T|Dias|600|;"
    tots = tots & "S|txtAux(6)|T|Resto|600|;"
    tots = tots & "S|txtAux(4)|T|FP.Ext|700|;N||||0|;N||||0|;N||||0|;N||||0|;N||||0|;N||||0|;N||||0|;"
'    tots = tots & "S|txtAux(4)|T|Alvic|600|;"
'    tots = tots & "S|txtAux(7)|T|Cod.Ext|1000|;"
'    tots = tots & "S|txtAux(8)|T|Socio|900|;"
'    tots = tots & "S|btnBuscar(1)|B|||;S|txtAux2(8)|T|Nombre|1100|;"
'    tots = tots & "N||||0|;S|Combo1(4)|C|FPV|500|;"
'    tots = tots & "S|txtAux(2)|T|Cta.Contable|1200|;"
    
    ' ### [Monica] 12/09/2006
    ' añadido para mostrar el nombre de la cuenta de contabilidad en caso de haya contabilidad
    If vParamAplic.NumeroConta <> 0 Then
'        tots = tots & "S|btnBuscar(0)|B|||;S|txtAux2(2)|T|Nombre de Cuenta|3000|;"
    End If
    
    arregla tots, DataGrid1, Me
    
    CargaForaGrid
    
    DataGrid1.ScrollBars = dbgAutomatic
    DataGrid1.Columns(0).Alignment = dbgRight
'   DataGrid1.Columns(2).Alignment = dbgRight
End Sub


Private Sub CargaForaGrid()
Dim I As Integer
'Dim tipclien
    On Error Resume Next

    '#############
    ' DAVID
    '    Este BUF habria que estudiarlo.
    ' Realmente ocurre que cuando hace en ponerforagridgeneral, hay un momento que
    ' el column no esta establecido
    ' If DataGrid1.Columns.Count >= 1 Then



    If DataGrid1.Columns.Count > 2 Then
        txtAux(7).Text = DataGrid1.Columns(14).Text
        txtAux(8).Text = DataGrid1.Columns(15).Text
        txtAux2(8).Text = DataGrid1.Columns(16).Text
        
        PosicionarCombo Me.Combo1(4), DataGrid1.Columns(17).Text

        txtAux(2).Text = DataGrid1.Columns(19).Text
        
        If vParamAplic.NumeroConta <> 0 Then
            txtAux2(2).Text = DataGrid1.Columns(20).Text
        End If
    End If

End Sub



Private Sub txtAux_GotFocus(Index As Integer)
    ConseguirFocoLin txtAux(Index)
End Sub


Private Sub txtAux_LostFocus(Index As Integer)
Dim SQL As String

    If Not PerderFocoGnral(txtAux(Index), Modo) Then Exit Sub
    
    Select Case Index
        Case 0, 3, 4, 5, 6
            PonerFormatoEntero txtAux(Index)
        Case 1
            txtAux(Index).Text = UCase(txtAux(Index).Text)
        
        Case 2 'cuenta contable
            If txtAux(Index).Text = "" Then Exit Sub
            txtAux2(Index).Text = PonerNombreCuenta(txtAux(Index), Modo)
            
        Case 8 ' codigo de socio
            If PonerFormatoEntero(txtAux(Index)) Then
                txtAux2(Index).Text = PonerNombreDeCod(txtAux(Index), "ssocio", "nomsocio", "codsocio", "N")
            Else
                txtAux2(Index).Text = ""
            End If
    End Select
    
End Sub

Private Function DatosOk() As Boolean
'Dim Datos As String
Dim b As Boolean
Dim SQL As String
Dim Mens As String


    b = CompForm(Me)
    If Not b Then Exit Function
    
    If Modo = 3 Then   'Estamos insertando
         If ExisteCP(txtAux(0)) Then b = False
    End If
    
    'Comprobaciones de TESORERIA
    If Modo = 4 Then
        'Estoy modificando
        
       '[Monica]25/04/2018: solo en el caso de que hayan cambiado algo de la contailidad
        If TipFPAnt <> Combo1(0).ListIndex Or NroVtoAnt <> txtAux(5).Text Or DiasVtoAnt <> txtAux(3).Text Or RestoVtoAnt <> txtAux(6).Text Then
            If Not PuedeModificarFPenContab Then b = False
        End If
    End If
    
    DatosOk = b
End Function

Private Sub CargaCombo()
Dim Cad As String
Dim I As Byte
Dim Rs As ADODB.Recordset
Dim SQL As String

    On Error GoTo ErrCarga
    For I = 0 To Combo1.Count - 1
        Combo1(I).Clear
    Next I
    
'[Monica]26/06/2013: cambio los valores fijos por los valores leyendolos de la contabilidad
'    Combo1(0).AddItem "Efectivo"
'    Combo1(0).ItemData(Combo1(0).NewIndex) = 0
'    Combo1(0).AddItem "Transferencia"
'    Combo1(0).ItemData(Combo1(0).NewIndex) = 1
'    Combo1(0).AddItem "Talon"
'    Combo1(0).ItemData(Combo1(0).NewIndex) = 2
'    Combo1(0).AddItem "Pagare"
'    Combo1(0).ItemData(Combo1(0).NewIndex) = 3
'    Combo1(0).AddItem "Recibo bancario"
'    Combo1(0).ItemData(Combo1(0).NewIndex) = 4
'    Combo1(0).AddItem "Confirming"
'    Combo1(0).ItemData(Combo1(0).NewIndex) = 5
'    '[Monica]04/01/2013: Efectivos
'    Combo1(0).AddItem "Tarjeta"
'    Combo1(0).ItemData(Combo1(0).NewIndex) = 6
    
    '[Monica]26/06/2013: Modifico la carga del tipo de forma de pago haciendola coincidir con la de la conta
    If vParamAplic.ContabilidadNueva Then
        SQL = "select * from tipofpago order by 2"
    Else
        SQL = "select * from stipoformapago order by 2"
    End If
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, ConnConta, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    While Not Rs.EOF
        Combo1(0).AddItem DBLet(Rs.Fields(1).Value, "T")
        Combo1(0).ItemData(Combo1(0).NewIndex) = DBLet(Rs.Fields(0).Value, "N")
    
        Rs.MoveNext
    Wend
    
    Set Rs = Nothing
    
    For I = 1 To 3
        Combo1(I).AddItem "No"
        Combo1(I).ItemData(Combo1(I).NewIndex) = 0
        Combo1(I).AddItem "Si"
        Combo1(I).ItemData(Combo1(I).NewIndex) = 1
    Next I
    
    Combo1(4).AddItem "Normal"
    Combo1(4).ItemData(Combo1(4).NewIndex) = 0
    Combo1(4).AddItem "Vales Cajeados"
    Combo1(4).ItemData(Combo1(4).NewIndex) = 1
    Combo1(4).AddItem "Devoluciones Pendientes"
    Combo1(4).ItemData(Combo1(4).NewIndex) = 2
    
    Exit Sub
    
ErrCarga:
    If Err.Number <> 0 Then MuestraError Err.Number, "Cargar datos combo.", Err.Description
End Sub

Private Sub PonerOpcionesMenu()
    PonerOpcionesMenuGeneral Me
    PonerOpcionesMenuGeneralNew Me
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

Private Sub printNou()
    With frmImprimir2
        .cadTabla2 = "sforpa"
        .Informe2 = "rManFpago.rpt"
        If CadB <> "" Then
            '.cadRegSelec = Replace(SQL2SF(CadB), "clientes", "clientes_1")
            .cadRegSelec = SQL2SF(CadB)
        Else
            .cadRegSelec = ""
        End If
        ' *** repasar el nom de l'adodc ***
        '.cadRegActua = Replace(POS2SF(Data1, Me), "clientes", "clientes_1")
        .cadRegActua = POS2SF(adodc1, Me)
        ' *** repasar codEmpre ***
        .cadTodosReg = ""
        '.cadTodosReg = "{itinerar.codempre} = " & codEmpre
        ' *** repasar si li pose ordre o no ****
        .OtrosParametros2 = "pEmpresa='" & vEmpresa.nomEmpre & "'|pOrden={sforpa.codforpa}|"
        '.OtrosParametros2 = "pEmpresa='" & vEmpresa.NomEmpre & "'|"
        ' *** posar el nº de paràmetres que he posat en OtrosParametros2 ***
        '.NumeroParametros2 = 1
        .NumeroParametros2 = 2
        ' ******************************************************************
        .MostrarTree2 = False
        .InfConta2 = False
        .ConSubInforme2 = False

        .Show vbModal
    End With
End Sub

'' ### [DavidV] 26/04/2006: Activar/desactivar la rueda del ratón.
'Private Sub DataGrid1_GotFocus()
'  WheelHook DataGrid1
'End Sub
'Private Sub DataGrid1_Lostfocus()
'  WheelUnHook
'End Sub

'Private Sub txtAux_KeyPress(Index As Integer, KeyAscii As Integer)
'    KEYpress KeyAscii
'End Sub
Private Sub txtAux_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = teclaBuscar Then
        If Modo = 1 Or Modo = 3 Or Modo = 4 Then
            Select Case Index
                Case 2: KEYBusqueda KeyAscii, 0 'cuenta contable
            End Select
        End If
    Else
        KEYpress KeyAscii
    End If
End Sub

Private Sub txtAux_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'Alvançar/Retrocedir els camps en les fleches de desplaçament del teclat.
    KEYdown KeyCode
End Sub

Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, Modo, cerrar
    If cerrar Then Unload Me
End Sub

Private Sub KEYBusqueda(KeyAscii As Integer, indice As Integer)
    KeyAscii = 0
    imgBuscar_Click (indice)
End Sub



Private Function InsertarRegistro() As Boolean
Dim MenError As String
Dim bol As Boolean, Existe As Boolean
Dim cambiaSQL As Boolean
Dim devuelve As String
Dim vSQL As String
Dim SQL As String

    On Error GoTo EInsertar
    
    bol = True
    
    vSQL = CadenaInsertarDesdeForm(Me)
    
    'Aqui empieza transaccion
    Conn.BeginTrans
    MenError = "Error al insertar en la tabla Formas de Pago (forpago)."
    Conn.Execute vSQL, , adCmdText
    
    ConnConta.BeginTrans
    
    If vParamAplic.ContabilidadNueva Then
        SQL = DevuelveDesdeBDNew(cConta, "formapago", "nomforpa", "codforpa", txtAux(0).Text, "N")
        If SQL = "" Then
            SQL = "insert into formapago (codforpa, nomforpa, tipforpa, numerove, primerve, restoven) values (" & DBSet(txtAux(0).Text, "N") & ","
            SQL = SQL & DBSet(txtAux(1).Text, "T") & "," & DBSet(Combo1(0).ItemData(Combo1(0).ListIndex), "N") & ","
            SQL = SQL & DBSet(txtAux(5).Text, "N") & "," & DBSet(txtAux(3).Text, "N") & "," & DBSet(txtAux(6).Text, "N") & ")"
            
            ConnConta.Execute SQL
        Else
            SQL = "update formapago set nomforpa = " & DBSet(txtAux(1).Text, "T") & ", tipforpa = " & DBSet(Combo1(0).ItemData(Combo1(0).ListIndex), "N")
            SQL = SQL & ", numerove = " & DBSet(txtAux(5).Text, "N")
            SQL = SQL & ", primerve = " & DBSet(txtAux(3).Text, "N")
            SQL = SQL & ", restoven = " & DBSet(txtAux(6).Text, "N")
            SQL = SQL & " where codforpa = " & DBSet(txtAux(0).Text, "N")
            
            ConnConta.Execute SQL
        End If
    
    Else
        SQL = DevuelveDesdeBDNew(cConta, "sforpa", "nomforpa", "codforpa", txtAux(0).Text, "N")
        If SQL = "" Then
            SQL = "insert into sforpa (codforpa, nomforpa, tipforpa) values (" & DBSet(txtAux(0).Text, "N") & ","
            SQL = SQL & DBSet(txtAux(1).Text, "T") & "," & DBSet(Combo1(0).ItemData(Combo1(0).ListIndex), "N") & ")"
            
            ConnConta.Execute SQL
        Else
            SQL = "update sforpa set nomforpa = " & DBSet(txtAux(1).Text, "T") & ", tipforpa = " & DBSet(Combo1(0).ItemData(Combo1(0).ListIndex), "N")
            SQL = SQL & " where codforpa = " & DBSet(txtAux(0).Text, "N")
            
            ConnConta.Execute SQL
        End If
    End If
    
EInsertar:
    If Err.Number <> 0 Then
        MenError = "Insertando Forma de Pago." & vbCrLf & "----------------------------" & vbCrLf & MenError
        MuestraError Err.Number, MenError, Err.Description
        bol = False
    End If
    If bol Then
        Conn.CommitTrans
        ConnConta.CommitTrans
        InsertarRegistro = True
    Else
        Conn.RollbackTrans
        ConnConta.RollbackTrans
        InsertarRegistro = False
    End If
End Function

Private Function ModificaRegistro() As Boolean
Dim b As Boolean
Dim MenError As String
Dim SQL As String
Dim vWhere As String

    On Error GoTo EModificarCab

    Conn.BeginTrans
    ConnConta.BeginTrans
    
    b = ModificaDesdeFormulario(Me)
    If b Then
        '[Monica]25/04/2018: solo en el caso de que hayan cambiado algo de la contailidad
        If TipFPAnt <> Combo1(0).ListIndex Or NroVtoAnt <> txtAux(5).Text Or DiasVtoAnt <> txtAux(3).Text Or RestoVtoAnt <> txtAux(6).Text Then
            If vParamAplic.ContabilidadNueva Then
                SQL = "update formapago set nomforpa = " & DBSet(txtAux(1).Text, "T") & ", tipforpa = " & DBSet(Combo1(0).ItemData(Combo1(0).ListIndex), "N")
                SQL = SQL & ", numerove = " & DBSet(txtAux(5).Text, "N")
                SQL = SQL & ", primerve = " & DBSet(txtAux(3).Text, "N")
                SQL = SQL & ", restoven = " & DBSet(txtAux(6).Text, "N")
                SQL = SQL & " where codforpa = " & DBSet(txtAux(0).Text, "N")
            Else
                SQL = "update sforpa set nomforpa = " & DBSet(txtAux(1).Text, "T") & ", tipforpa = " & DBSet(Combo1(0).ItemData(Combo1(0).ListIndex), "N")
                SQL = SQL & " where codforpa = " & DBSet(txtAux(0).Text, "N")
            End If
            ConnConta.Execute SQL
        Else
            SQL = "update formapago set nomforpa = " & DBSet(txtAux(1).Text, "T")
            SQL = SQL & " where codforpa = " & DBSet(txtAux(0).Text, "N")
            ConnConta.Execute SQL
        End If
    End If
    

EModificarCab:
    If Err.Number <> 0 Then
        MenError = "Modificando Forma de Pago." & vbCrLf & "----------------------------" & vbCrLf & MenError
        MuestraError Err.Number, MenError, Err.Description
        b = False
    End If
    If b Then
        ModificaRegistro = True
        Conn.CommitTrans
        ConnConta.CommitTrans
    Else
        ModificaRegistro = False
        Conn.RollbackTrans
        ConnConta.RollbackTrans
    End If
End Function



