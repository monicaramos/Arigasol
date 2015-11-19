VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmBusLinFactu 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Líneas de Factura"
   ClientHeight    =   8220
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14085
   Icon            =   "frmBusLinFactu.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8220
   ScaleWidth      =   14085
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnBuscar 
      Appearance      =   0  'Flat
      Caption         =   "+"
      Height          =   300
      Index           =   2
      Left            =   3900
      MaskColor       =   &H00000000&
      TabIndex        =   22
      ToolTipText     =   "Buscar Artículo"
      Top             =   3240
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.CommandButton btnBuscar 
      Appearance      =   0  'Flat
      Caption         =   "+"
      Height          =   300
      Index           =   1
      Left            =   1680
      MaskColor       =   &H00000000&
      TabIndex        =   21
      ToolTipText     =   "Buscar Artículo"
      Top             =   3240
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.TextBox txtAux2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   285
      Index           =   0
      Left            =   6480
      TabIndex        =   20
      Top             =   3240
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.TextBox txtAux 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   290
      Index           =   12
      Left            =   10560
      MaxLength       =   15
      TabIndex        =   12
      Tag             =   "Importe|N|N|0|9999999999.99|slhfac|implinea|##,###,##0.00||"
      Text            =   "Importe"
      Top             =   3240
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   290
      Index           =   11
      Left            =   1200
      MaxLength       =   10
      TabIndex        =   2
      Tag             =   "Fecha Factura|F|N|||slhfac|fecfactu|dd/mm/yyyy|S|"
      Text            =   "fecfactu"
      Top             =   3240
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txtAux 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   290
      Index           =   2
      Left            =   600
      MaxLength       =   7
      TabIndex        =   1
      Tag             =   "Nº Factura|N|N|0|9999999|slhfac|numfactu|0000000|S|"
      Text            =   "Fac"
      Top             =   3240
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txtAux 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   290
      Index           =   3
      Left            =   1800
      MaxLength       =   2
      TabIndex        =   3
      Tag             =   "Número de línea|N|N|1|9999|slhfac|numlinea|0000|S|"
      Text            =   "li"
      Top             =   3240
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox txtAux 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   290
      Index           =   1
      Left            =   240
      MaxLength       =   3
      TabIndex        =   0
      Tag             =   "Letra Serie|T|N|||slhfac|letraser||S|"
      Text            =   "L"
      Top             =   3240
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox txtAux 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   290
      Index           =   9
      Left            =   5760
      MaxLength       =   6
      TabIndex        =   9
      Tag             =   "Artículo|N|N|0|999999|slhfac|codartic|000000||"
      Text            =   "Arti"
      Top             =   3240
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   290
      Index           =   4
      Left            =   2160
      MaxLength       =   8
      TabIndex        =   4
      Tag             =   "Albaran|T|N|||slhfac|numalbar|||"
      Text            =   "Alb"
      Top             =   3240
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton btnBuscar 
      Appearance      =   0  'Flat
      Caption         =   "+"
      Height          =   300
      Index           =   0
      Left            =   6240
      MaskColor       =   &H00000000&
      TabIndex        =   19
      ToolTipText     =   "Buscar Artículo"
      Top             =   3240
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   290
      Index           =   5
      Left            =   3120
      MaxLength       =   10
      TabIndex        =   5
      Tag             =   "Fecha Albaran|F|N|||slhfac|fecalbar|dd/mm/yyyy||"
      Text            =   "Fec"
      Top             =   3240
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   290
      Index           =   6
      Left            =   4080
      MaxLength       =   5
      TabIndex        =   6
      Tag             =   "Hora|FHH|N|||slhfac|horalbar|hh:mm:ss||"
      Text            =   "Hor"
      Top             =   3240
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txtAux 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   290
      Index           =   7
      Left            =   4680
      MaxLength       =   1
      TabIndex        =   7
      Tag             =   "Turno|N|N|0|9|slhfac|codturno|0||"
      Text            =   "T"
      Top             =   3240
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox txtAux 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   290
      Index           =   8
      Left            =   5040
      MaxLength       =   8
      TabIndex        =   8
      Tag             =   "Tarjeta|N|N|0|99999999|slhfac|numtarje|00000000||"
      Text            =   "tar"
      Top             =   3240
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox txtAux 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   290
      Index           =   10
      Left            =   8760
      MaxLength       =   10
      TabIndex        =   10
      Tag             =   "Cantidad|N|N|0|99999999.99|slhfac|cantidad|#,##0.00||"
      Text            =   "Can"
      Top             =   3240
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox txtAux 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   290
      Index           =   0
      Left            =   9720
      MaxLength       =   10
      TabIndex        =   11
      Tag             =   "Precio|N|N|0|99999999.999|slhfac|preciove|#,##0.000||"
      Text            =   "Pre"
      Top             =   3240
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   11220
      TabIndex        =   13
      Top             =   7470
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   12510
      TabIndex        =   15
      Top             =   7470
      Visible         =   0   'False
      Width           =   1095
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmBusLinFactu.frx":000C
      Height          =   6750
      Left            =   120
      TabIndex        =   17
      Top             =   540
      Width           =   13785
      _ExtentX        =   24315
      _ExtentY        =   11906
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
   Begin VB.CommandButton cmdRegresar 
      Caption         =   "&Regresar"
      Height          =   375
      Left            =   12480
      TabIndex        =   18
      Top             =   7440
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Height          =   555
      Index           =   1
      Left            =   120
      TabIndex        =   14
      Top             =   7440
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
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Width           =   2175
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
      TabIndex        =   23
      Top             =   0
      Width           =   14085
      _ExtentX        =   24844
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   19
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Buscar"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Ver todos"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Object.Tag             =   "2"
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Modificar"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Object.Tag             =   "2"
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Imprimir"
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Primero"
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Anterior"
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Siguiente"
         EndProperty
         BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Último"
         EndProperty
      EndProperty
      BorderStyle     =   1
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
         Enabled         =   0   'False
         HelpContextID   =   2
         Shortcut        =   ^N
         Visible         =   0   'False
      End
      Begin VB.Menu mnModificar 
         Caption         =   "&Modificar"
         Enabled         =   0   'False
         HelpContextID   =   2
         Shortcut        =   ^M
         Visible         =   0   'False
      End
      Begin VB.Menu mnEliminar 
         Caption         =   "&Eliminar"
         Enabled         =   0   'False
         HelpContextID   =   2
         Shortcut        =   ^E
         Visible         =   0   'False
      End
      Begin VB.Menu mnBarra2 
         Caption         =   ""
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu mnSalir 
         Caption         =   "&Salir"
         Shortcut        =   ^S
      End
   End
End
Attribute VB_Name = "frmBusLinFactu"
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
' 6. En la funció posamaxlength() repasar el maxlength de TextAux(0)
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
' ********************************************************************************

Option Explicit

Public DatosADevolverBusqueda As String    'Tendra el nº de text que quiere que devuelva, empipados
Public Event DatoSeleccionado(CadenaSeleccion As String)

'codi per al registe que s'afegix al cridar des d'atre formulari.
'Obrir en modo Insertar i tornar datos del registre insertat
Public NuevoCodigo As String

'codigo que tiene el campo en el momento que se llama desde otro formulario
'nos situamos en ese valor
Public CodigoActual As String
Dim btnPrimero As Byte 'Variable que indica el nº del Botó PrimerRegistro en la Toolbar1

Private WithEvents frmFpa As frmManFpago
Attribute frmFpa.VB_VarHelpID = -1
Private WithEvents frmcli As frmManClien
Attribute frmcli.VB_VarHelpID = -1
Private WithEvents frmArt As frmManArtic
Attribute frmArt.VB_VarHelpID = -1
Private WithEvents frmTra As frmManTraba
Attribute frmTra.VB_VarHelpID = -1
Private WithEvents frmB As frmBuscaGrid
Attribute frmB.VB_VarHelpID = -1
Private WithEvents frmC As frmCal
Attribute frmC.VB_VarHelpID = -1

Private CadenaConsulta As String
Private CadB As String
Private PrimeraVez As Boolean
Private HaDevueltoDatos As Boolean


Dim tipoF As String
Dim Modo As Byte
Dim CampoSeleccionado As Integer


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
Dim i As Byte

    On Error Resume Next
    
    Modo = vModo
'    PonerIndicador lblIndicador, Modo
    b = (Modo = 2)
    If b Then
        PonerContRegIndicador
    Else
        PonerIndicador lblIndicador, Modo
    End If
    
    For i = 0 To txtAux.Count - 1 'els txtAux del grid
        txtAux(i).visible = Not b
    Next i
    btnBuscar(0).visible = Not b
    btnBuscar(1).visible = Not b
    btnBuscar(2).visible = Not b
    txtAux2(0).visible = Not b
    
    cmdAceptar.visible = Not b
    cmdCancelar.visible = Not b
    DataGrid1.Enabled = b
    
    'Si es regresar
    If DatosADevolverBusqueda <> "" Then cmdRegresar.visible = b
    
    If Modo = 3 Or Modo = 4 Or Modo = 1 Then i = 4 'Insertar/Modificar o busqueda
    BloquearImgBuscar Me, i
    
    PonerLongCampos
    PonerModoOpcionesMenu 'Activar/Desact botones de menu segun Modo
    PonerOpcionesMenu 'Activar/Desact botones de menu segun permisos del usuario
    
    'Si estamos modo Modificar bloquear clave primaria
    BloquearTxt txtAux(0), (Modo = 4)
    
    PonerFocoGrid Me.DataGrid1
    
    If Err.Number <> 0 Then Err.Clear
End Sub

Private Sub PonerModoOpcionesMenu()
Dim b As Boolean
On Error Resume Next

    b = (Modo = 2)
    'Busqueda
    Toolbar1.Buttons(2).Enabled = b
    Me.mnBuscar.Enabled = b
    'Ver Todos
    Toolbar1.Buttons(3).Enabled = b
    Me.mnVerTodos.Enabled = b
    
    Toolbar1.Buttons(6).Enabled = True
    Me.mnSalir.Enabled = True
    If Err.Number <> 0 Then Err.Clear
End Sub

Private Sub BotonAnyadir()
    Dim NumF As String
    Dim anc As Single
    Dim i As Integer
    
'   ' ### [Monica] 21/09/2006
'   ' cuando añado se carga todo sql grid estaba la instruccion de abajo
    CargaGrid CadB  'primer de tot carregue tot el grid
'    CargaGrid "codclave = -1" 'primer de tot carregue tot el grid
   
    CadB = ""
    '******************** canviar taula i camp **************************
    If (DatosADevolverBusqueda <> "") And NuevoCodigo <> "" Then
        NumF = NuevoCodigo
    Else
        NumF = SugerirCodigoSiguienteStr("scaalb", "codclave")
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
    
    LLamaLineas anc, 3 '(limpia los campos)
    
    txtAux(0).Text = NumF
 '   FormateaCampo txtAux(0)
    For i = 1 To 10
        txtAux(i).Text = ""
    Next i
    txtAux2(5).Text = ""
    txtAux2(7).Text = ""
    txtAux(2).Text = Format(Now, "dd/mm/yyyy") ' Fecha x defecto
    txtAux(3).Text = Format(Now, "hh:mm") ' Hora x defecto
       
    'Ponemos el foco
    If (DatosADevolverBusqueda <> "") And NuevoCodigo <> "" Then
        PonerFoco txtAux(1)
    Else
        PonerFoco txtAux(1)
    End If
End Sub

Private Sub BotonVerTodos()
    CadB = ""
    CargaGrid ""
    PonerModo 2
End Sub

Private Sub BotonBuscar()
Dim i As Integer
    ' ***************** canviar per la clau primaria ********
    CargaGrid "numfactu = -1"
    '*******************************************************************************
    'Buscar
    For i = 0 To txtAux.Count - 1
        txtAux(i).Text = ""
    Next i

    LLamaLineas DataGrid1.Top + 216, 1
    PonerFoco txtAux(1)
End Sub


Private Sub LLamaLineas(alto As Single, xModo As Byte)
Dim i As Byte

    DeseleccionaGrid Me.DataGrid1
    PonerModo xModo
    
    'Fijamos el ancho
'    txtAux(0).Top = alto - 20
    For i = 0 To txtAux.Count - 1
        txtAux(i).Top = alto
    Next i
    For i = 0 To btnBuscar.Count - 1
        btnBuscar(i).Top = alto
    Next i
    txtAux2(0).Top = alto
End Sub


Private Sub PonerLongCampos()
'Modificar el MaxLength del campo en funcion de si es modo de búsqueda o no
'para los campos que permitan introducir criterios más largos del tamaño del campo
'Si maxlength=6 para codprove, en modo busqueda Maxlenth=13 'el doble + 1
    
    'para los txtAux
    PonerLongCamposGnral Me, Modo, 3
End Sub

Private Sub btnBuscar_Click(Index As Integer)
    TerminaBloquear
    Select Case Index
        Case 0 'Articulo
            Set frmArt = New frmManArtic
            frmArt.DatosADevolverBusqueda = "0|1|"
            frmArt.CodigoActual = txtAux(7).Text
            frmArt.Show vbModal
            Set frmArt = Nothing
            PonerFoco txtAux(7)
        Case 1, 2 ' Fecha
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
            If Index = 1 Then
                If txtAux(11).Text <> "" Then frmC.NovaData = txtAux(11).Text
                CampoSeleccionado = 11
            Else
                If txtAux(5).Text <> "" Then frmC.NovaData = txtAux(5).Text
                CampoSeleccionado = 5
            End If
            
            ' ********************************************
        
            frmC.Show vbModal
            Set frmC = Nothing
            ' *** repasar si el camp es txtAux o Text1 ***
            If Index = 1 Then
                PonerFoco txtAux(11) '<===
            Else
                PonerFoco txtAux(5) '<===
            End If
            ' ********************************************
        
            
    End Select
    If Modo = 4 Then BLOQUEADesdeFormulario2 Me, adodc1, 1
End Sub

Private Sub cmdAceptar_Click()
Dim i As Integer

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
                If InsertarDesdeForm2(Me, 1) Then
                    CargaGrid
                    If (DatosADevolverBusqueda <> "") And NuevoCodigo <> "" Then
                        cmdCancelar_Click
    '                    If Not adodc1.Recordset.EOF Then adodc1.Recordset.MoveLast
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
                If ModificaDesdeFormulario2(Me) Then
                    TerminaBloquear
                    i = adodc1.Recordset.Fields(0)
                    PonerModo 2
                    CargaGrid CadB
'                    If CadB <> "" Then
'                        CargaGrid CadB
'                        lblIndicador.Caption = "BUSQUEDA: " & PonerContRegistros(Me.adodc1)
'                    Else
'                        CargaGrid
'                        lblIndicador.Caption = ""
'                    End If
                    adodc1.Recordset.Find (adodc1.Recordset.Fields(0).Name & " =" & i)
                    PonerFocoGrid Me.DataGrid1
                    
                    'si se ha modificado la empresa que estamos conectados
                    'refrescar los datos de la clase
'                    If Val(vEmpresa.codEmpre) = Val(txtAux(0).Text) Then
'                       If vEmpresa.LeerDatos(vEmpresa.codEmpre) = False Then
'                            MsgBox "No se han podido cargar los datos de la empresa.", vbExclamation
'                            AccionesCerrar
'                            End
'                       End If
'                    End If
                End If
            End If
    End Select
End Sub

Private Sub cmdCancelar_Click()
On Error Resume Next

    Select Case Modo
        Case 1 'BUSQUEDA
            CargaGrid CadB
'        Case 3 'INSERTAR
'            DataGrid1.AllowAddNew = False
'            If Not adodc1.Recordset.EOF Then adodc1.Recordset.MoveFirst
'        Case 4 'MODIFICAR
'            TerminaBloquear
    End Select
    
    If Not adodc1.Recordset.EOF Then
'        CargaForaGrid
    Else
        LimpiarCampos
    End If
    
    PonerModo 2
'    If CadB <> "" Then
'        lblIndicador.Caption = "BUSQUEDA: " & PonerContRegistros(Me.adodc1)
'    Else
'        lblIndicador.Caption = ""
'    End If
    
    
    PonerFocoGrid Me.DataGrid1
    
    If Err.Number <> 0 Then Err.Clear
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
    If Not Me.adodc1.Recordset.EOF Then
        frmHcoFact.letraserie = Me.adodc1.Recordset.Fields(0).Value
        frmHcoFact.numfactu = Me.adodc1.Recordset.Fields(1).Value
        frmHcoFact.Tipo = 0 ' seleccionamos la tabla schfac
        frmHcoFact.Show vbModal
    End If
    
End Sub

Private Sub DataGrid1_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub DataGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
Dim i As Byte
    
    If Modo <> 4 Then 'Modificar
'        CargaForaGrid
    Else 'vamos a Insertar
        For i = 0 To txtAux.Count - 1
            txtAux(i).Text = ""
        Next i
    End If
    
'    If (Modo = 2 Or Modo = 0) Then
'        If CadB = "" Then
'            lblIndicador.Caption = PonerContRegistros(Me.adodc1)
'        Else
'            lblIndicador.Caption = "BUSQUEDA: " & PonerContRegistros(Me.adodc1)
'        End If
'    End If
    PonerContRegIndicador
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
                SituarData Me.adodc1, "codempre=" & CodigoActual, "", True
            End If
        End If
    End If
End Sub

Private Sub Form_Load()
Dim i As Integer

    PrimeraVez = True
    
    ' ICONITOS DE LA BARRA
    btnPrimero = 16 'index del botó "primero"
    With Me.Toolbar1
        .HotImageList = frmPpal.imgListComun_OM
        .DisabledImageList = frmPpal.imgListComun_BN
        .ImageList = frmPpal.imgListComun
        'el 1 i el 2 son separadors
        .Buttons(3).Image = 1   'Buscar
        .Buttons(4).Image = 2   'Todos
        'el 5 i el 6 son separadors
        .Buttons(7).Image = 3   'Insertar
        .Buttons(8).Image = 4   'Modificar
        .Buttons(9).Image = 5   'Borrar
        'el 10 i el 11 son separadors
        .Buttons(12).Image = 10  'Imprimir
        .Buttons(13).Image = 11  'Salir
        'el 14 i el 15 son separadors
        .Buttons(btnPrimero).Image = 6  'Primero
        .Buttons(btnPrimero + 1).Image = 7 'Anterior
        .Buttons(btnPrimero + 2).Image = 8 'Siguiente
        .Buttons(btnPrimero + 3).Image = 9 'Último
    End With


    '## A mano
'    chkVistaPrevia.Value = CheckValueLeer(Name)

    '****************** canviar la consulta *********************************+
    CadenaConsulta = "SELECT letraser,numfactu,fecfactu,numlinea,numalbar,fecalbar,"
    CadenaConsulta = CadenaConsulta & "horalbar, codturno, numtarje, slhfac.codartic, sartic.nomartic, "
    CadenaConsulta = CadenaConsulta & "cantidad, preciove, implinea "
    CadenaConsulta = CadenaConsulta & " FROM slhfac, sartic "
    CadenaConsulta = CadenaConsulta & " WHERE slhfac.codartic = sartic.codartic "
    '************************************************************************
    
    CadB = ""
    CargaGrid "numfactu  = -1"
'    CargaForaGrid
   
'    If (DatosADevolverBusqueda <> "") And NuevoCodigo <> "" Then
'        BotonAnyadir
'    Else
'        PonerModo 2
'    End If
End Sub


Private Sub frmB_Selecionado(CadenaDevuelta As String)
   txtAux(6).Text = RecuperaValor(CadenaDevuelta, 1)
   txtAux(6).Text = Format(txtAux(6).Text, "00000000")
End Sub

Private Sub frmC_Selec(vFecha As Date)
    ' *** repasar si el camp es txtAux o Text1 ***
    txtAux(CampoSeleccionado).Text = Format(vFecha, "dd/mm/yyyy") '<===
    ' ********************************************
End Sub

Private Sub frmArt_DatoSeleccionado(CadenaSeleccion As String)
    txtAux(9).Text = RecuperaValor(CadenaSeleccion, 1)
    FormateaCampo txtAux(7)
    txtAux2(0).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub imgBuscar_Click(Index As Integer)
    TerminaBloquear
    Select Case Index
        Case 0 'Trabajador
            Set frmTra = New frmManTraba
            frmTra.DatosADevolverBusqueda = "0|1|"
            frmTra.CodigoActual = txtAux(11).Text
            frmTra.Show vbModal
            Set frmTra = Nothing
            PonerFoco txtAux(11)
            
        Case 1 'F.Pago
            Set frmFpa = New frmManFpago
            frmFpa.DatosADevolverBusqueda = "0|1|"
            frmFpa.CodigoActual = txtAux(12).Text
            frmFpa.Show vbModal
            Set frmFpa = Nothing
            PonerFoco txtAux(12)
            
    End Select
    
    If Modo = 4 Then BLOQUEADesdeFormulario2 Me, adodc1, 1
End Sub

Private Sub frmTra_DatoSeleccionado(CadenaSeleccion As String)
    txtAux(11).Text = RecuperaValor(CadenaSeleccion, 1) 'codtraba
    FormateaCampo txtAux(11)
    txtAux2(11).Text = RecuperaValor(CadenaSeleccion, 2) 'nomtraba
End Sub

Private Sub frmFpa_DatoSeleccionado(CadenaSeleccion As String)
    txtAux(12).Text = RecuperaValor(CadenaSeleccion, 1) 'cod fpa
    FormateaCampo txtAux(12)
    txtAux2(12).Text = RecuperaValor(CadenaSeleccion, 2) 'nomforpa
    
End Sub

Private Sub mnBuscar_Click()
    BotonBuscar
End Sub

Private Sub mnSalir_Click()
    Screen.MousePointer = vbDefault
    Unload Me
End Sub

Private Sub mnVerTodos_Click()
    BotonVerTodos
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 3
                mnBuscar_Click
        Case 4
                mnVerTodos_Click
        Case 13
                mnSalir_Click
    End Select
End Sub

Private Sub CargaGrid(Optional vSQL As String)
    Dim sql As String
    Dim tots As String
    
'    adodc1.ConnectionString = Conn
    If vSQL <> "" Then
        sql = CadenaConsulta & " And  " & vSQL
    Else
        sql = CadenaConsulta
    End If
    '********************* canviar el ORDER BY *********************++
    sql = sql & " ORDER BY numfactu, numlinea"
    '**************************************************************++
    
    CargaGridGnral Me.DataGrid1, Me.adodc1, sql, PrimeraVez
    
    tots = "S|txtAux(1)|T|Serie|800|;S|txtAux(2)|T|Factura|800|;S|txtAux(11)|T|F.Factura|1200|;S|btnBuscar(1)|B|||;N||||0|;"
    tots = tots & "S|txtAux(4)|T|Albaran|900|;S|txtAux(5)|T|Fecha|1200|;S|btnBuscar(2)|B|||;S|txtAux(6)|T|Hora|750|;"
    tots = tots & "S|txtAux(7)|T|Tur.|400|;S|txtAux(8)|T|Tarjeta|900|;S|txtAux(9)|T|Articulo|800|;"
    tots = tots & "S|btnBuscar(0)|B|||;S|txtAux2(0)|T|Denominación|2180|;S|txtAux(10)|T|Cantidad|1100|;"
    tots = tots & "S|txtAux(0)|T|Precio|1000|;S|txtAux(12)|T|Importe|1200|;"
    arregla tots, DataGrid1, Me
    
    DataGrid1.ScrollBars = dbgAutomatic
      
    If Not adodc1.Recordset.EOF Then
'        CargaForaGrid
    Else
        LimpiarCampos
    End If
    
    DataGrid1.Columns(0).Alignment = dbgRight
'    DataGrid1.Columns(2).Alignment = dbgRight
      
'   'Habilitamos modificar y eliminar
'   Toolbar1.Buttons(7).Enabled = Not adodc1.Recordset.EOF
'   Toolbar1.Buttons(8).Enabled = Not adodc1.Recordset.EOF
'   mnModificar.Enabled = Not adodc1.Recordset.EOF
'   mnEliminar.Enabled = Not adodc1.Recordset.EOF
   
End Sub

Private Sub txtAux_GotFocus(Index As Integer)
    ConseguirFocoLin txtAux(Index)
End Sub

Private Sub txtAux_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    KEYdown KeyCode
End Sub

Private Sub txtAux_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = teclaBuscar Then
        If Modo = 1 Or Modo = 3 Or Modo = 4 Then
            Select Case Index
                Case 5: 'fecha de albaran
                    KeyAscii = 0
                    btnBuscar_Click (2)
                Case 9: 'articulo
                    KeyAscii = 0
                    btnBuscar_Click (0)
                Case 11: 'fecha de factura
                    KeyAscii = 0
                    btnBuscar_Click (1)
                
                Case 12: KEYBusqueda KeyAscii, 1 'F.Pago
            End Select
        End If
    Else
        KEYpress KeyAscii
    End If
End Sub

Private Sub txtAux_LostFocus(Index As Integer)
Dim cadMen As String
Dim Famia As String

    'If Modo = 1 Then Exit Sub 'Busquedas
    If Not PerderFocoGnral(txtAux(Index), Modo) Then Exit Sub
    
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub
    
    Select Case Index
        Case 1 ' letra de serie
            txtAux(Index).Text = UCase(txtAux(Index).Text)
        
        Case 2 ' numfactu
            PonerFormatoEntero txtAux(Index)
            
        Case 1, 13 'ALBARAN , MATRICULA
            txtAux(Index).Text = UCase(txtAux(Index).Text)
        
        Case 11, 5 'FECHA
            PonerFormatoFecha txtAux(Index)
        
        Case 6 'Hora
            PonerFormatoHora txtAux(Index)
        
        Case 9 'cod articulo
            If PonerFormatoEntero(txtAux(Index)) Then
                txtAux2(0).Text = PonerNombreDeCod(txtAux(Index), "sartic", "nomartic", "codartic", "N")
                If txtAux2(0).Text = "" Then
                    cadMen = "No existe el Articulo: " & txtAux(Index).Text & vbCrLf
                    MsgBox cadMen, vbExclamation
                End If
            End If
            
        Case 10 'CANTIDAD
            If Modo = 1 Then Exit Sub
            If txtAux(Index).Text <> "" Then PonerFormatoDecimal txtAux(Index), 2
           
        Case 0 'PRECIO
' ### [Monica] 25/09/2006
' he quitado las dos lineas siguientes y he puesto ponerformatodecimal
'            cadMen = TransformaPuntosComas(txtAux(Index).Text)
'            txtAux(Index).Text = Format(cadMen, "##,##0.000")
            If Modo = 1 Then Exit Sub
            If txtAux(Index).Text <> "" Then PonerFormatoDecimal txtAux(Index), 2
            
        Case 12 'IMPORTE
            If Modo = 1 Then Exit Sub
            If txtAux(Index).Text <> "" Then PonerFormatoDecimal txtAux(Index), 3
            
    End Select
End Sub

Private Function DatosOk() As Boolean
Dim b As Boolean
Dim sql As String

    b = CompForm(Me)
    If Not b Then Exit Function
    
    If Modo = 3 Then   'Estamos insertando
        'comprobar si ya existe el campo de clave primaria
        If ExisteCP(txtAux(0)) Then b = False
        
    End If
    
    ' comprobamos que la tarjeta introducida esta asociada al socio
    sql = ""
    sql = DevuelveDesdeBDNew(cPTours, "starje", "numtarje", "codsocio", txtAux(5).Text, "N", , "numtarje", txtAux(6).Text, "N")
    If sql = "" Then
        MsgBox "El número de Tarjeta introducida no corresponde al Socio. Revise.", vbExclamation
        PonerFoco txtAux(6)
        b = False
    End If
    
    DatosOk = b
End Function

Private Sub PonerOpcionesMenu()
    PonerOpcionesMenuGeneral Me
    PonerOpcionesMenuGeneralNew Me
End Sub

Private Sub LimpiarCampos()
Dim i As Integer

    On Error Resume Next

    txtAux2(0).Text = ""
    
    If Err.Number <> 0 Then Err.Clear
End Sub

Private Sub BotonImprimir()
        frmAlbTurno.Show vbModal
End Sub

Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, 0, cerrar
    If cerrar Then Unload Me
End Sub

Private Sub KEYBusqueda(KeyAscii As Integer, indice As Integer)
    KeyAscii = 0
    imgBuscar_Click (indice)
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

' ### [DavidV] 26/04/2006: Activar/desactivar la rueda del ratón.
Private Sub DataGrid1_GotFocus()
  WheelHook DataGrid1
End Sub
Private Sub DataGrid1_Lostfocus()
  WheelUnHook
End Sub

