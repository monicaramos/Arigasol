VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmAlbaran 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Albaranes"
   ClientHeight    =   8220
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15570
   Icon            =   "frmAlbaran.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8220
   ScaleWidth      =   15570
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtAux 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   17
      Left            =   3960
      MaxLength       =   3
      TabIndex        =   17
      Tag             =   "Colectivo|N|S|||ssocio|codcoope|000||"
      Top             =   7530
      Width           =   735
   End
   Begin VB.TextBox txtAux2 
      BackColor       =   &H80000018&
      Enabled         =   0   'False
      Height          =   285
      Index           =   17
      Left            =   4800
      TabIndex        =   39
      Top             =   7530
      Width           =   2775
   End
   Begin VB.TextBox txtAux 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   16
      Left            =   13110
      MaxLength       =   15
      TabIndex        =   16
      Tag             =   "Importe Vale|N|S|||scaalb|importevale|##,##0.00||"
      Top             =   6960
      Width           =   1575
   End
   Begin VB.TextBox txtAux 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   15
      Left            =   11280
      MaxLength       =   15
      TabIndex        =   15
      Tag             =   "Dto Alvic|N|S|||scaalb|dtoalvic|##,##0.000||"
      Top             =   6960
      Width           =   1575
   End
   Begin VB.TextBox txtAux 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   14
      Left            =   9450
      MaxLength       =   11
      TabIndex        =   14
      Tag             =   "Kilómetros|N|S|||scaalb|kilometros|###,###,##0||"
      Top             =   6960
      Width           =   1575
   End
   Begin VB.CommandButton btnBuscar 
      Appearance      =   0  'Flat
      Caption         =   "+"
      Height          =   290
      Index           =   2
      Left            =   1860
      MaskColor       =   &H00000000&
      TabIndex        =   35
      ToolTipText     =   "Buscar Fecha"
      Top             =   2760
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.TextBox txtAux 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   290
      Index           =   10
      Left            =   9900
      MaxLength       =   17
      TabIndex        =   10
      Tag             =   "Importe|N|N|||scaalb|importel|##,##0.00||"
      Text            =   "Imp"
      Top             =   2760
      Width           =   555
   End
   Begin VB.TextBox txtAux 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   290
      Index           =   9
      Left            =   9300
      MaxLength       =   15
      TabIndex        =   9
      Tag             =   "Precio|N|N||99999.999|scaalb|preciove|##,##0.000||"
      Text            =   "Precio"
      Top             =   2760
      Width           =   555
   End
   Begin VB.TextBox txtAux 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   290
      Index           =   8
      Left            =   8640
      MaxLength       =   15
      TabIndex        =   8
      Tag             =   "Cantidad|N|N||99999.999|scaalb|cantidad|##,##0.000||"
      Text            =   "Can"
      Top             =   2760
      Width           =   555
   End
   Begin VB.CommandButton btnBuscar 
      Appearance      =   0  'Flat
      Caption         =   "+"
      Height          =   290
      Index           =   1
      Left            =   7200
      MaskColor       =   &H00000000&
      TabIndex        =   34
      ToolTipText     =   "Buscar Articulo"
      Top             =   2760
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.TextBox txtAux 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   290
      Index           =   7
      Left            =   6600
      MaxLength       =   6
      TabIndex        =   7
      Tag             =   "Articulo|N|N|0|999999|scaalb|codartic|000000|N|"
      Text            =   "Art"
      Top             =   2760
      Width           =   555
   End
   Begin VB.TextBox txtAux2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   285
      Index           =   7
      Left            =   7440
      TabIndex        =   33
      Top             =   2760
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox txtAux 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   290
      Index           =   6
      Left            =   5760
      MaxLength       =   13
      TabIndex        =   6
      Tag             =   "Tarjeta|N|N|0|9999999999999|scaalb|numtarje|0000000000000||"
      Text            =   "Tarjeta"
      Top             =   2760
      Width           =   795
   End
   Begin VB.TextBox txtAux 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   290
      Index           =   4
      Left            =   2640
      MaxLength       =   1
      TabIndex        =   4
      Tag             =   "Turno|N|N|0|9|scaalb|codturno|0||"
      Text            =   "Tu"
      Top             =   2760
      Width           =   315
   End
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   290
      Index           =   3
      Left            =   2040
      MaxLength       =   8
      TabIndex        =   3
      Tag             =   "Hora|FHH|N|||scaalb|horalbar|hh:mm:ss||"
      Text            =   "hora"
      Top             =   2760
      Width           =   555
   End
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   290
      Index           =   2
      Left            =   1440
      MaxLength       =   10
      TabIndex        =   2
      Tag             =   "Fecha|F|N|||scaalb|fecalbar|dd/mm/yyyy||"
      Top             =   2760
      Width           =   555
   End
   Begin VB.TextBox txtAux 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   290
      Index           =   0
      Left            =   180
      MaxLength       =   6
      TabIndex        =   0
      Tag             =   "Clave|N|N|1|999999|scaalb|codclave|000000|S|"
      Top             =   2760
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.TextBox txtAux2 
      BackColor       =   &H80000018&
      Enabled         =   0   'False
      Height          =   285
      Index           =   12
      Left            =   4800
      TabIndex        =   23
      Top             =   6960
      Width           =   2775
   End
   Begin VB.TextBox txtAux 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   12
      Left            =   3960
      MaxLength       =   2
      TabIndex        =   12
      Tag             =   "F.Pago|N|N|0|99|scaalb|codforpa|00||"
      Top             =   6960
      Width           =   735
   End
   Begin VB.TextBox txtAux2 
      BackColor       =   &H80000018&
      Enabled         =   0   'False
      Height          =   285
      Index           =   11
      Left            =   960
      TabIndex        =   22
      Top             =   6960
      Width           =   2775
   End
   Begin VB.TextBox txtAux 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   11
      Left            =   120
      MaxLength       =   4
      TabIndex        =   11
      Tag             =   "Trabajador|N|N|0|9999|scaalb|codtraba|0000||"
      Top             =   6960
      Width           =   735
   End
   Begin VB.TextBox txtAux 
      Height          =   285
      Index           =   13
      Left            =   7680
      MaxLength       =   10
      TabIndex        =   13
      Tag             =   "Matrícula|T|S|||scaalb|matricul|||"
      Top             =   6960
      Width           =   1575
   End
   Begin VB.TextBox txtAux2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   285
      Index           =   5
      Left            =   3840
      TabIndex        =   21
      Top             =   2760
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.TextBox txtAux 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   290
      Index           =   5
      Left            =   3030
      MaxLength       =   6
      TabIndex        =   5
      Tag             =   "Cliente|N|N|0|999999|scaalb|codsocio|000000|N|"
      Text            =   "Cli"
      Top             =   2760
      Width           =   555
   End
   Begin VB.CommandButton btnBuscar 
      Appearance      =   0  'Flat
      Caption         =   "+"
      Height          =   290
      Index           =   0
      Left            =   3600
      MaskColor       =   &H00000000&
      TabIndex        =   20
      ToolTipText     =   "Buscar Cliente"
      Top             =   2760
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   13060
      TabIndex        =   18
      Top             =   7500
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   14320
      TabIndex        =   19
      Top             =   7500
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   290
      Index           =   1
      Left            =   780
      MaxLength       =   8
      TabIndex        =   1
      Tag             =   "Albaran|T|N|||scaalb|numalbar|||"
      Text            =   "Alb"
      Top             =   2760
      Width           =   555
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmAlbaran.frx":000C
      Height          =   5910
      Left            =   120
      TabIndex        =   26
      Top             =   540
      Width           =   15285
      _ExtentX        =   26961
      _ExtentY        =   10425
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
      Left            =   14320
      TabIndex        =   32
      Top             =   7500
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Height          =   555
      Index           =   1
      Left            =   120
      TabIndex        =   24
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
         TabIndex        =   25
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
      TabIndex        =   27
      Top             =   0
      Width           =   15570
      _ExtentX        =   27464
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   14
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
            Object.ToolTipText     =   "Borrar Turno"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Total selección"
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Cargas sin Tarjeta"
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Imprimir"
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Imprimir Ticket"
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin VB.CheckBox chkVistaPrevia 
         Caption         =   "Vista previa"
         Height          =   195
         Left            =   8400
         TabIndex        =   28
         Top             =   120
         Visible         =   0   'False
         Width           =   1215
      End
   End
   Begin VB.Image imgBuscar 
      Height          =   240
      Index           =   2
      Left            =   5430
      MousePointer    =   4  'Icon
      ToolTipText     =   "Buscar F.Pago"
      Top             =   7290
      Width           =   240
   End
   Begin VB.Label Label1 
      Caption         =   "Colectivo del cliente"
      Height          =   255
      Index           =   5
      Left            =   3960
      TabIndex        =   40
      Top             =   7320
      Width           =   1425
   End
   Begin VB.Label Label1 
      Caption         =   "Descuento Vale"
      Height          =   255
      Index           =   4
      Left            =   13110
      TabIndex        =   38
      Top             =   6720
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Descuento Alvic"
      Height          =   255
      Index           =   3
      Left            =   11280
      TabIndex        =   37
      Top             =   6720
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Kilómetros"
      Height          =   255
      Index           =   2
      Left            =   9480
      TabIndex        =   36
      Top             =   6720
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "F.Pago"
      Height          =   255
      Index           =   1
      Left            =   3960
      TabIndex        =   31
      Top             =   6720
      Width           =   780
   End
   Begin VB.Image imgBuscar 
      Height          =   240
      Index           =   1
      Left            =   4800
      MousePointer    =   4  'Icon
      ToolTipText     =   "Buscar F.Pago"
      Top             =   6720
      Width           =   240
   End
   Begin VB.Label Label1 
      Caption         =   "Trabajador"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   30
      Top             =   6720
      Width           =   795
   End
   Begin VB.Image imgBuscar 
      Height          =   240
      Index           =   0
      Left            =   960
      MousePointer    =   4  'Icon
      ToolTipText     =   "Buscar Trabajador"
      Top             =   6720
      Width           =   240
   End
   Begin VB.Label Label1 
      Caption         =   "Matrícula"
      Height          =   255
      Index           =   7
      Left            =   7680
      TabIndex        =   29
      Top             =   6720
      Width           =   1215
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
      Begin VB.Menu mnBorrarTurno 
         Caption         =   "&Borrar Turno"
         HelpContextID   =   2
         Shortcut        =   ^U
      End
      Begin VB.Menu mnTotalSeleccion 
         Caption         =   "&Total Seleccion"
         Shortcut        =   ^V
      End
      Begin VB.Menu mnCargasSinTarjeta 
         Caption         =   "&Cargas sin Tarjeta"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnImprimir 
         Caption         =   "&Imprimir"
         Shortcut        =   ^I
      End
      Begin VB.Menu mnImprimirTicket 
         Caption         =   "Imprimir &Ticket"
         Shortcut        =   ^T
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
Attribute VB_Name = "frmAlbaran"
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

Private WithEvents frmFPa As frmManFpago
Attribute frmFPa.VB_VarHelpID = -1
Private WithEvents frmcli As frmManClien
Attribute frmcli.VB_VarHelpID = -1
Private WithEvents frmArt As frmManArtic
Attribute frmArt.VB_VarHelpID = -1
Private WithEvents frmTra As frmManTraba
Attribute frmTra.VB_VarHelpID = -1
Private WithEvents frmB As frmBuscaGrid
Attribute frmB.VB_VarHelpID = -1
Private WithEvents frmB1 As frmBuscaGrid
Attribute frmB1.VB_VarHelpID = -1
Private WithEvents frmCol As frmManCoope
Attribute frmCol.VB_VarHelpID = -1
Private WithEvents frmC As frmCal
Attribute frmC.VB_VarHelpID = -1

Private CadenaConsulta As String
Private CadB As String
Private PrimeraVez As Boolean
Private HaDevueltoDatos As Boolean

Dim ValorAnt As String
Dim SocioAnt As String

Dim tipoF As String
Dim Modo As Byte

Dim CodTipoMov As String
Dim vCont As CContador



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
Dim I As Byte

    On Error Resume Next
    
    Modo = vModo
'    PonerIndicador lblIndicador, Modo
    b = (Modo = 2)
    If b Then
        PonerContRegIndicador
    Else
        PonerIndicador lblIndicador, Modo
    End If
    
    For I = 1 To 10 'els txtAux del grid
        txtAux(I).visible = Not b
    Next I
    btnBuscar(0).visible = Not b
    btnBuscar(1).visible = Not b
    btnBuscar(2).visible = Not b
    txtAux2(5).visible = Not b
    txtAux2(7).visible = Not b
    
    txtAux(1).Enabled = (Modo = 1 And vParamAplic.Cooperativa = 4) Or (vParamAplic.Cooperativa <> 4)
    
    For I = 11 To 16
        BloquearTxt txtAux(I), b
    Next I
       
    cmdAceptar.visible = Not b
    cmdCancelar.visible = Not b
    DataGrid1.Enabled = b
    
    'Si es regresar
    If DatosADevolverBusqueda <> "" Then cmdRegresar.visible = b
    
    If Modo = 3 Or Modo = 4 Or Modo = 1 Then I = 4 'Insertar/Modificar o busqueda
    BloquearImgBuscar Me, I
    
    PonerLongCampos
    PonerModoOpcionesMenu 'Activar/Desact botones de menu segun Modo
    PonerOpcionesMenu 'Activar/Desact botones de menu segun permisos del usuario
    
    ' Si estamos modo Modificar bloquear clave primaria
    BloquearTxt txtAux(0), (Modo = 4)
    
    '[Monica]07/04/2017: el colectivo sólo se puede consultar, no se puede añadir ni modificar pq está en la tabla de socios
    BloquearTxt txtAux(17), Not (Modo = 1)
    
    imgBuscar(2).visible = (Modo = 1)
    imgBuscar(2).Enabled = (Modo = 1)
    
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
    
    'Insertar
    Toolbar1.Buttons(6).Enabled = b 'And Not DeConsulta
    Me.mnNuevo.Enabled = b 'And Not DeConsulta
    
    b = (b And adodc1.Recordset.RecordCount > 0) 'And Not DeConsulta
    'Modificar
    Toolbar1.Buttons(7).Enabled = b
    Me.mnModificar.Enabled = b
    'Eliminar
    Toolbar1.Buttons(8).Enabled = b
    Me.mnEliminar.Enabled = b
    
    
    'Borrar turno
    '[Monica]25/11/2010: en Pobla del Duc introducen los albaranes a mano no permitimos borrar turno
    Toolbar1.Buttons(9).Enabled = b And (vParamAplic.Cooperativa <> 4)
    Me.mnBorrarTurno.Enabled = b And (vParamAplic.Cooperativa <> 4)
    
    'Total seleccion
    Toolbar1.Buttons(10).Enabled = b
    Me.mnTotalSeleccion.Enabled = b
    'Imprimir
    Toolbar1.Buttons(13).Enabled = b
    Me.mnImprimirTicket.Enabled = b
    
    If Err.Number <> 0 Then Err.Clear
End Sub

Private Sub BotonAnyadir()
    Dim NumF As String
    Dim anc As Single
    Dim I As Integer
    
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
    
    txtAux(0).Text = ""
    If vParamAplic.Cooperativa <> 4 Then
        txtAux(0).Text = NumF
    End If
 '   FormateaCampo txtAux(0)
    For I = 1 To 10
        txtAux(I).Text = ""
    Next I
    txtAux2(5).Text = ""
    txtAux2(7).Text = ""
    txtAux(2).Text = Format(Now, "dd/mm/yyyy") ' Fecha x defecto
    txtAux(3).Text = Format(Now, "hh:mm:ss") ' Hora x defecto
       
    'Ponemos el foco
    If (DatosADevolverBusqueda <> "") And NuevoCodigo <> "" Then
        PonerFoco txtAux(1)
    Else
        PonerFoco txtAux(1)
    End If
    
    If vParamAplic.Cooperativa = 4 Then
        txtAux(1).Text = "" 'NumF
        PonerFoco txtAux(2)
    End If
    
    CadenaCambio = NumF
    
End Sub

Private Sub BotonVerTodos()
    CadB = ""
    CargaGrid ""
    PonerModo 2
End Sub

Private Sub BotonBuscar()
Dim I As Integer
    ' ***************** canviar per la clau primaria ********
    CargaGrid "codclave = -1"
    '*******************************************************************************
    'Buscar
    For I = 0 To txtAux.Count - 1
        txtAux(I).Text = ""
    Next I
    txtAux2(5).Text = ""
    txtAux2(7).Text = ""
    
    '[Monica]07/04/2017: campo colectivo
    txtAux2(17).Text = ""

    LLamaLineas DataGrid1.Top + 216, 1
    PonerFoco txtAux(1)
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
        anc = DataGrid1.RowTop(DataGrid1.Row) + 545
    End If

    'Llamamos al form
    For I = 0 To 5
        txtAux(I).Text = DataGrid1.Columns(I).Text
    Next I
    
    For I = 6 To 7
        txtAux(I).Text = DataGrid1.Columns(I + 1).Text
    Next I
    
    For I = 8 To 10
        txtAux(I).Text = DataGrid1.Columns(I + 2).Text
    Next I
    
    txtAux2(5).Text = DataGrid1.Columns(6).Text
    txtAux2(7).Text = DataGrid1.Columns(9).Text

' ### [Monica] 18/12/2006
    CargarValoresAnteriores Me

    LLamaLineas anc, 4
   
    'Como es modificar
    
    '02/03/2007 añadidas esta linea para dar aviso si cambian el socio de que la FP no se corresponde
    SocioAnt = txtAux(5).Text
    
    If vParamAplic.Cooperativa = 4 Then
        PonerFoco txtAux(2)
    Else
        PonerFoco txtAux(1)
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub LLamaLineas(alto As Single, xModo As Byte)
Dim I As Byte

    DeseleccionaGrid Me.DataGrid1
    PonerModo xModo
    
    'Fijamos el ancho
'    txtAux(0).Top = alto - 20
    For I = 0 To 10
        txtAux(I).Top = alto
    Next I
    btnBuscar(0).Top = alto '- 10
    btnBuscar(1).Top = alto '- 10
    btnBuscar(2).Top = alto
    txtAux2(5).Top = alto
    txtAux2(7).Top = alto
    
'    If (Modo = 1) Or (Modo = 3) Then 'Busqueda/Insertar
'        For i = 0 To txtAux.Count - 1
'            txtAux(i).Text = ""
'        Next i
'        txtAux2(11).Text = ""
'        txtAux2(12).Text = ""
'    End If
End Sub

Private Sub BotonEliminar()
Dim SQL As String
Dim temp As Boolean

    On Error GoTo Error2
    
    'Ciertas comprobaciones
    If adodc1.Recordset.EOF Then Exit Sub
'    If Not SepuedeBorrar Then Exit Sub
    
    'El registre de codi 0 no es pot Modificar ni Eliminar
    ' ### [Monica] 27/10/2006 he quitado la linea de no eliminar el codigo 0
'    If EsCodigoCero(CStr(adodc1.Recordset.Fields(0).Value), FormatoCampo(txtAux(0))) Then Exit Sub
    
    '*************** canviar els noms i el DELETE **********************************
    SQL = "¿Seguro que desea eliminar el Albaran?"
    SQL = SQL & vbCrLf & "Albaran: " & adodc1.Recordset.Fields(1)
    
    If MsgBox(SQL, vbQuestion + vbYesNo) = vbYes Then
        'Hay que eliminar
        NumRegElim = adodc1.Recordset.AbsolutePosition
        If Not EliminarLinea Then Exit Sub
        CadenaCambio = SQL
        InsertarCambios "scaalb", ValorNulo, adodc1.Recordset.Fields(1)
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


Private Sub BotonCargasSinTarjeta()
Dim SQL As String
Dim sql2 As String
Dim Sql3 As String
Dim Sql4 As String
Dim Rs As ADODB.Recordset
Dim Tarjeta As String
Dim temp As Boolean

    On Error GoTo Error2
    
    SQL = "select scaalb.codclave, scaalb.codsocio, scaalb.numtarje "
    SQL = SQL & " from scaalb"
    SQL = SQL & " where not (codsocio,numtarje) in (select codsocio, numtarje from starje)"
    
    Set Rs = New ADODB.Recordset ' Crear objeto
    Rs.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText ' abrir cursor
    
    If Rs.EOF Then
        MsgBox "No hay cargas con tarjetas inexistentes." & vbCrLf & vbCrLf, vbInformation
        Rs.Close
        Set Rs = Nothing
        Exit Sub
    End If
    
    ' almacenamos las claves cuyos socios tengan mas de una tarjeta
    Sql4 = ""
    
    Screen.MousePointer = vbHourglass
    
    Conn.BeginTrans
    
    While Not Rs.EOF
        sql2 = "select count(*) from starje where codsocio = " & DBSet(Rs!codsocio, "N")
        
        If TotalRegistros(sql2) = 1 Then
            Tarjeta = ""
            Tarjeta = DevuelveDesdeBDNew(cPTours, "starje", "numtarje", "codsocio", Rs!codsocio, "N")
            
            Sql3 = "update scaalb set numtarje = " & DBSet(Tarjeta, "N") & " where codclave = " & DBSet(Rs!Codclave, "N")
            
            Conn.Execute Sql3
        
        Else
            Sql4 = Sql4 & DBSet(Rs!Codclave, "N") & ","
        
        End If
        
        Rs.MoveNext
    Wend
    
    Rs.Close
    Set Rs = Nothing
    
    
Error2:
    If Err.Number <> 0 Then
        Screen.MousePointer = vbDefault

        MuestraError Err.Number, "Actualizando Tarjetas", Err.Description
        Conn.RollbackTrans
    Else
        Screen.MousePointer = vbDefault
        
        Conn.CommitTrans
        
        ' quitamos la ultima coma
        Sql4 = Mid(Sql4, 1, Len(Sql4) - 1)
        Sql4 = "(" & Sql4 & ")"
        
        ' mostramos unicamente los registros que no hemos podido modificar
        CadB = "scaalb.codclave in " & Sql4
        CargaGrid CadB
        PonerModoOpcionesMenu
    End If
    
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
        Case 0 'Cliente
            Set frmcli = New frmManClien
            frmcli.DatosADevolverBusqueda = "0|1|"
            frmcli.CodigoActual = txtAux(5).Text
            frmcli.Show vbModal
            Set frmcli = Nothing
            PonerFoco txtAux(5)
        Case 1 'Articulo
            Set frmArt = New frmManArtic
            frmArt.DatosADevolverBusqueda = "0|1|"
            frmArt.CodigoActual = txtAux(7).Text
            frmArt.Show vbModal
            Set frmArt = Nothing
            PonerFoco txtAux(7)
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
            If txtAux(2).Text <> "" Then frmC.NovaData = txtAux(2).Text
            ' ********************************************
        
            frmC.Show vbModal
            Set frmC = Nothing
            ' *** repasar si el camp es txtAux o Text1 ***
            PonerFoco txtAux(2) '<===
            ' ********************************************
        
            
    End Select
    If Modo = 4 Then BLOQUEADesdeFormulario2 Me, adodc1, 1
End Sub

Private Sub cmdAceptar_Click()
Dim I As Long

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
'[Monica]16/11/2010: LLevamos control de stock en ventas
'                If InsertarDesdeForm2(Me, 1) Then
                 If InsertarLinea(0) Then
                    InsertarCambios "scaalb", ValorNulo, txtAux(1).Text
                    '[Monica]25/11/2010: si hay impresora de tickets se lanza la impresion
                    If vParamAplic.Cooperativa = 4 And vParamAplic.ImpresoraTickets <> "" Then
                        mnImprimirTicket_Click
                    End If
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
'[Monica]16/11/2010: actualizamos stock
'                If ModificaDesdeFormulario2(Me) Then
                If ModificarLinea Then
                    InsertarCambios "scaalb", ValorAnterior, txtAux(1).Text
                    TerminaBloquear
                    I = adodc1.Recordset.Fields(0).Value
                    
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
        Case 3 'INSERTAR
            DataGrid1.AllowAddNew = False
            If Not adodc1.Recordset.EOF Then adodc1.Recordset.MoveFirst
        Case 4 'MODIFICAR
            TerminaBloquear
    End Select
    
    If Not adodc1.Recordset.EOF Then
        CargaForaGrid
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

Private Sub DataGrid1_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub DataGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
Dim I As Byte
    
    If Modo <> 4 Then 'Modificar
        CargaForaGrid
    Else 'vamos a Insertar
        For I = 0 To txtAux.Count - 1
            txtAux(I).Text = ""
        Next I
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

Private Sub CargaForaGrid()
Dim I As Integer
'Dim tipclien
    On Error Resume Next

    If DataGrid1.Columns.Count > 2 Then
        txtAux(11).Text = DataGrid1.Columns(13).Text
        txtAux(12).Text = DataGrid1.Columns(15).Text
        txtAux(13).Text = DataGrid1.Columns(17).Text
        '[Monica]24/06/2013: introducimos los kilometros
        txtAux(14).Text = Format(DataGrid1.Columns(18).Text, "###,###,##0")
        '[Monica]24/08/2015: falta introducir el dto alvic
        txtAux(15).Text = Format(DataGrid1.Columns(19).Text, "###,##0.000")
        'vale de descuento
        txtAux(16).Text = Format(DataGrid1.Columns(20).Text, "###,##0.00")
        
        txtAux2(11).Text = PonerNombreDeCod(txtAux(11), "straba", "nomtraba", "codtraba", "N")
        txtAux2(12).Text = PonerNombreDeCod(txtAux(12), "sforpa", "nomforpa", "codforpa", "N")
        
        '[Monica]07/04/2017: visualizamos el colectivo del socio
        txtAux(17).Text = Format(DataGrid1.Columns(21).Text, "000")
        txtAux2(17).Text = PonerNombreDeCod(txtAux(17), "scoope", "nomcoope", "codcoope", "N")
        
    End If
End Sub

Private Sub DataGrid1_DblClick()
    If cmdRegresar.visible Then cmdRegresar_Click
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
Dim I As Integer

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
        .Buttons(9).Image = 20   'Eliminar Turno
        .Buttons(10).Image = 21   'Totales
        .Buttons(11).Image = 25   'Cargas sin tarjeta
        .Buttons(12).Image = 10  'Imprimir
        .Buttons(13).Image = 26 'Imprimir Ticket
        .Buttons(14).Image = 11  'Salir
    End With

    'cargar IMAGES de busqueda
    For I = 0 To Me.imgBuscar.Count - 1
        Me.imgBuscar(I).Picture = frmPpal.imgListImages16.ListImages(1).Picture
    Next I

    '## A mano
'    chkVistaPrevia.Value = CheckValueLeer(Name)

    '****************** canviar la consulta *********************************+
    CadenaConsulta = "Select scaalb.codclave, scaalb.numalbar, scaalb.fecalbar, "
    CadenaConsulta = CadenaConsulta & "scaalb.horalbar, scaalb.codturno, scaalb.codsocio, "
    CadenaConsulta = CadenaConsulta & "ssocio.nomsocio, scaalb.numtarje, scaalb.codartic, "
    CadenaConsulta = CadenaConsulta & "sartic.nomartic, scaalb.cantidad, scaalb.preciove, scaalb.importel, "
    CadenaConsulta = CadenaConsulta & "scaalb.codtraba, straba.nomtraba, scaalb.codforpa, "
    CadenaConsulta = CadenaConsulta & "sforpa.nomforpa, scaalb.matricul, scaalb.kilometros, "
    '[Monica]24/08/2015: faltaba introducir el dto alvic
    CadenaConsulta = CadenaConsulta & "scaalb.dtoalvic, scaalb.importevale, "
    '[Monica]07/04/2017: podemos buscar por campo codcoope
    CadenaConsulta = CadenaConsulta & "ssocio.codcoope "
    
    CadenaConsulta = CadenaConsulta & "from ((((scaalb INNER JOIN ssocio ON scaalb.codsocio=ssocio.codsocio) "
    CadenaConsulta = CadenaConsulta & " INNER JOIN sartic ON scaalb.codartic=sartic.codartic) "
    CadenaConsulta = CadenaConsulta & " INNER JOIN straba ON scaalb.codtraba=straba.codtraba) "
    CadenaConsulta = CadenaConsulta & " INNER JOIN sforpa ON scaalb.codforpa=sforpa.codforpa) "
    '************************************************************************
    
    
    CodTipoMov = "ALV"
    
    If vParamAplic.Cooperativa = 4 Then
        txtAux(0).Tag = "Clave|N|S|1|999999|scaalb|codclave|000000|S|"
        txtAux(1).Tag = "Albaran|T|S|||scaalb|numalbar|||"
    Else
        '[Monica]02/04/2013: añadido, en el caso de castelduc el numero de albaran es varchar, TICKET
        If vParamAplic.Cooperativa = 3 Then
            txtAux(0).Tag = "Clave|N|N|1|999999|scaalb|codclave|000000|S|"
            txtAux(1).Tag = "Albaran|T|N|||scaalb|numalbar|0000000||"
        Else
            txtAux(0).Tag = "Clave|N|N|1|999999|scaalb|codclave|000000|S|"
            txtAux(1).Tag = "Albaran|N|N|||scaalb|numalbar|0000000||"
        End If
    End If
    CadB = ""
    CargaGrid "codclave = -1"
    CargaForaGrid
   
    '[Monica]24/08/2015: solo tienen descuento alvic en el caso de las facturas de regaixo
    Label1(3).visible = (vParamAplic.Cooperativa = 2)
    txtAux(15).Enabled = (vParamAplic.Cooperativa = 2)
    txtAux(15).visible = (vParamAplic.Cooperativa = 2)
   
    Label1(4).visible = (vParamAplic.Cooperativa = 2)
    txtAux(16).Enabled = (vParamAplic.Cooperativa = 2)
    txtAux(16).visible = (vParamAplic.Cooperativa = 2)
   
   
'    If (DatosADevolverBusqueda <> "") And NuevoCodigo <> "" Then
'        BotonAnyadir
'    Else
'        PonerModo 2
'    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
'    CheckValueGuardar Me.Name, Me.chkVistaPrevia.Value
    DesBloqueoManual ("BORTUR")
    Screen.MousePointer = vbDefault
End Sub

Private Sub frmB_Selecionado(CadenaDevuelta As String)
   txtAux(6).Text = RecuperaValor(CadenaDevuelta, 1)
   txtAux(6).Text = Format(txtAux(6).Text, "00000000")
End Sub

Private Sub frmB1_Selecionado(CadenaDevuelta As String)
   txtAux(13).Text = RecuperaValor(CadenaDevuelta, 1)
End Sub



Private Sub frmC_Selec(vFecha As Date)
    ' *** repasar si el camp es txtAux o Text1 ***
    txtAux(2).Text = Format(vFecha, "dd/mm/yyyy") '<===
    ' ********************************************
End Sub

Private Sub frmCli_DatoSeleccionado(CadenaSeleccion As String)
    txtAux(5).Text = RecuperaValor(CadenaSeleccion, 1)
    FormateaCampo txtAux(5)
    txtAux2(5).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmArt_DatoSeleccionado(CadenaSeleccion As String)
    txtAux(7).Text = RecuperaValor(CadenaSeleccion, 1)
    FormateaCampo txtAux(7)
    txtAux2(7).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmCol_DatoSeleccionado(CadenaSeleccion As String)
    txtAux(17).Text = RecuperaValor(CadenaSeleccion, 1)
    FormateaCampo txtAux(17)
    txtAux2(17).Text = RecuperaValor(CadenaSeleccion, 2)
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
            Set frmFPa = New frmManFpago
            frmFPa.DatosADevolverBusqueda = "0|1|"
            frmFPa.CodigoActual = txtAux(12).Text
            frmFPa.Show vbModal
            Set frmFPa = Nothing
            PonerFoco txtAux(12)
            
        Case 2 ' colectivo del socio
            Set frmCol = New frmManCoope
            frmCol.DatosADevolverBusqueda = "0|1|"
            frmCol.CodigoActual = txtAux(12).Text
            frmCol.Show vbModal
            Set frmCol = Nothing
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


Private Sub mnBorrarTurno_Click()
    DesBloqueoManual ("BORTUR")
    If Not BloqueoManual("BORTUR", "1") Then
        MsgBox "No se puede realizar el Borre de Turno. Hay otro usuario realizándolo.", vbExclamation
        Screen.MousePointer = vbDefault
    Else
        CadenaDesdeOtroForm = ""
        frmBorreTurno.Show vbModal
        If CadenaDesdeOtroForm <> "" Then CargaGrid "codclave = -1"
    End If
End Sub

Private Sub mnBuscar_Click()
    BotonBuscar
End Sub

Private Sub mnCargasSinTarjeta_Click()
    BotonCargasSinTarjeta
End Sub

Private Sub mnEliminar_Click()
    BotonEliminar
End Sub

Private Sub mnImprimir_Click()
    BotonImprimir
End Sub

Private Sub mnImprimirTicket_Click()
Dim NroCopias As String
Dim lin As String
Dim Directo As Boolean

    If adodc1.Recordset.EOF Then Exit Sub
    
    If vParamAplic.ImpresoraTickets = "" Then
        MsgBox "No tiene indicada en parámetros cual es la impresora de Tickets.", vbExclamation
        Exit Sub
    End If
    
    Dim indRPT As Byte 'Indica el tipo de Documento en la tabla "scryst"
    Dim nomDocu As String 'Nombre de Informe rpt de crystal
    Dim ImprimeDirecto As Integer
     
    indRPT = 6 'Ticket de Entrada
     
    If Not PonerParamRPT(indRPT, "", 1, nomDocu) Then Exit Sub
    'Nombre fichero .rpt a Imprimir
    ' he añadido estas dos lineas para que llame al rpt correspondiente
    
    ActivaTicket

    Directo = True
    
    
    If txtAux(0).Text = "" Then txtAux(0).Text = adodc1.Recordset.Fields(0)
    If txtAux(2).Text = "" Then txtAux(2).Text = Format(adodc1.Recordset.Fields(2), "dd/mm/yyyy")
    
    
    If Directo Then
        '-- Impresion directa
        ImprimirElTicketDirecto2 txtAux(0).Text, CDate(txtAux(2).Text), True
        If CLng(txtAux(12).Text) = 8 Then ImprimirElTicketDirecto2 txtAux(0).Text, CDate(txtAux(2).Text), True
    Else
        frmImprimir.NombreRPT = nomDocu
        
        With frmVisReport
            .FormulaSeleccion = "{scaalb.numalbar}=""" & adodc1.Recordset!numalbar & """"
            .SoloImprimir = True
            .OtrosParametros = ""
            .NumeroParametros = 1
            .MostrarTree = False
            .Informe = App.path & "\informes\" & nomDocu    ' "ValEntrada.rpt"
            .InfConta = False
            .ConSubInforme = False
            .SubInformeConta = ""
            .Opcion = 0
            .ExportarPDF = False
            .Show vbModal
        End With
        
    End If
    
    DesactivaTicket
'    Else
'        NroCopias = InputBox("Introduzca el Número de Copias:", "", , 5000, 4000)
'
'        If NroCopias = "" Then Exit Sub
'
'        ' imprimimos
'        If EsNumerico(NroCopias) Then
'            ' impresion directa por la printer
'            ' ImprimirEntradaDirectaPrinter Text1(0).Text, CInt(NroCopias)
'            ' impresion directa por LPT
'            ImprimirEntradaDirectaLPT text1(0).Text, CInt(NroCopias)
'        End If
'    End If
End Sub

Private Sub mnModificar_Click()
    'Comprobaciones
    '--------------
    If adodc1.Recordset.EOF Then Exit Sub
    If adodc1.Recordset.RecordCount < 1 Then Exit Sub
    
    'El registre de codi 0 no es pot Modificar ni Eliminar
    ' ### [Monica] 27/10/2006
    ' he quitado la linea de no poder eliminar ni modificar el registro 0
    'If EsCodigoCero(CStr(adodc1.Recordset.Fields(0).Value), FormatoCampo(txtAux(0))) Then Exit Sub
    
    
    'Preparar para modificar
    '-----------------------
    If BLOQUEADesdeFormulario2(Me, adodc1, 1) Then BotonModificar
End Sub

Private Sub mnNuevo_Click()
    BotonAnyadir
End Sub

Private Sub mnSalir_Click()
    Screen.MousePointer = vbDefault
    Unload Me
End Sub

Private Sub mnTotalSeleccion_Click()
    CalcularSumaPantalla
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
        Case 9
            mnBorrarTurno_Click
        Case 10
            mnTotalSeleccion_Click
        Case 11 ' cargas sin tarjeta
            mnCargasSinTarjeta_Click
        Case 12 'Imprimir
            mnImprimir_Click
        Case 13 ' Imprimir Ticket
            mnImprimirTicket_Click
        Case 14 'Salir
            mnSalir_Click
    End Select
End Sub

Private Sub CargaGrid(Optional vSQL As String)
    Dim SQL As String
    Dim tots As String
    
'    adodc1.ConnectionString = Conn
    If vSQL <> "" Then
        SQL = CadenaConsulta & " WHERE " & vSQL
    Else
        SQL = CadenaConsulta
    End If
    '********************* canviar el ORDER BY *********************++
'    Sql = Sql & " ORDER BY codclave"
    SQL = SQL & " ORDER BY fecalbar, horalbar, codturno"
     '**************************************************************++
    
    CargaGridGnral Me.DataGrid1, Me.adodc1, SQL, PrimeraVez
    
    tots = "N|||||;S|txtAux(1)|T|Albaran|800|;S|txtAux(2)|T|Fecha|1150|;S|btnBuscar(2)|B||195|;S|txtAux(3)|T|Hora|700|;"
    tots = tots & "S|txtAux(4)|T|Tu.|400|;S|txtAux(5)|T|Cliente|800|;S|btnBuscar(0)|B||195|;S|txtAux2(5)|T|Nombre|2700|;"
    tots = tots & "S|txtAux(6)|T|Tarjeta|1400|;S|txtAux(7)|T|Articulo|800|;S|btnBuscar(1)|B||195|;"
    tots = tots & "S|txtAux2(7)|T|Denominacion|2580|;S|txtAux(8)|T|Cantidad|1100|;"
    tots = tots & "S|txtAux(9)|T|Precio|1100|;S|txtAux(10)|T|Importe|1200|;"
    tots = tots & "N|||||;N|||||;N|||||;N|||||;N|||||;N|||||;N|||||;N|||||;"
    '[Monica]07/04/2017: no se incluye el campo de colectivo
    tots = tots & "N|||||;"
    
    arregla tots, DataGrid1, Me
    DataGrid1.ScrollBars = dbgAutomatic
      
    If Not adodc1.Recordset.EOF Then
        CargaForaGrid
    Else
        LimpiarCampos
    End If
    
    DataGrid1.Columns(0).Alignment = dbgRight
      
      
      
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
                Case 5: 'cliente
                    KeyAscii = 0
                    btnBuscar_Click (0)
                Case 7: 'articulo
                    KeyAscii = 0
                    btnBuscar_Click (1)
                Case 2: 'fecha de albaran
                    KeyAscii = 0
                    btnBuscar_Click (2)
                Case 11: KEYBusqueda KeyAscii, 0 'trabajador
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

    If Modo = 1 Then Exit Sub             'Busquedas

    If Not PerderFocoGnral(txtAux(Index), Modo) Then Exit Sub
    
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub
    
    Select Case Index
        Case 0 'codclave
            PonerFormatoEntero txtAux(Index)
            
        Case 1, 13 'ALBARAN , MATRICULA
            txtAux(Index).Text = UCase(txtAux(Index).Text)
        
        Case 2 'FECHA
            PonerFormatoFecha txtAux(Index)
        
        Case 3 'Hora
            PonerFormatoHora txtAux(Index)
        
        Case 5 'cod cliente
            If PonerFormatoEntero(txtAux(Index)) Then
                txtAux2(5).Text = PonerNombreDeCod(txtAux(Index), "ssocio", "nomsocio", "codsocio", "N")
                If txtAux2(5).Text = "" Then
                    cadMen = "No existe el Cliente: " & txtAux(Index).Text & vbCrLf
                    cadMen = cadMen & "¿Desea crearlo?" & vbCrLf
                    If MsgBox(cadMen, vbQuestion + vbYesNo) = vbYes Then
                        Set frmcli = New frmManClien
                        frmcli.DatosADevolverBusqueda = "0|1|"
                        frmcli.NuevoCodigo = txtAux(Index).Text
                        txtAux(Index).Text = ""
                        TerminaBloquear
                        frmcli.Show vbModal
                        Set frmcli = Nothing
                        If Modo = 4 Then BLOQUEADesdeFormulario2 Me, adodc1, 1
                    Else
                        txtAux(Index).Text = ""
                    End If
                    PonerFoco txtAux(Index)
                Else
                    ' ### [Monica] 10/10/2006
                    ' mostramos las tarjetas que tiene ese cliente
                    
                    MandaBusquedaTarjetas "codsocio = " & DBSet(txtAux(5).Text, "N")
                    If vParamAplic.Cooperativa = 4 Then
                        MandaBusquedaMatriculas "codsocio = " & DBSet(txtAux(5).Text, "N")
                    End If
                    
                    '[Monica]07/04/2017: traemos si insertamos el colectivo
                    If Modo = 3 Then
                        If txtAux(5).Text <> "" Then
                            txtAux(17).Text = DevuelveDesdeBDNew(cPTours, "ssocio", "codcoope", "codsocio", DBLet(txtAux(5).Text, "N"), "N")
                            txtAux2(17).Text = PonerNombreDeCod(txtAux(17), "scoope", "nomcoope", "codcoope", "N")
                        End If
                    End If
                    
                    ' ### [Monica] 08/09/2006
                    ' solo si estamos en modo insertar
                    If Modo = 3 Then
                         If txtAux(7).Text = "" Then Exit Sub
                         txtAux(9).Text = CargarPrecio(txtAux(7).Text, txtAux(5).Text)
                         txtAux(9).Text = Format(txtAux(9).Text, "##,##0.000")
                         Famia = DevuelveDesdeBD("codfamia", "sartic", "codartic", txtAux(7).Text, "N")
                         tipoF = DevuelveDesdeBD("tipfamia", "sfamia", "codfamia", Famia, "N")
                         If tipoF = "1" Then
                           If Modo = 3 Then BloquearTxt txtAux(8), True
                           If Modo = 3 Then BloquearTxt txtAux(10), False
                           PonerFoco txtAux(10)
                         Else
                           If Modo = 3 Then BloquearTxt txtAux(10), True
                           If Modo = 3 Then BloquearTxt txtAux(8), False
                           PonerFoco txtAux(8)
                         End If
                    End If
                End If
            Else
                txtAux2(5).Text = ""
            End If
            
        Case 7 'cod articulo
            If PonerFormatoEntero(txtAux(Index)) Then
                txtAux2(7).Text = PonerNombreDeCod(txtAux(Index), "sartic", "nomartic", "codartic", "N")
                If txtAux2(7).Text = "" Then
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
                Else
                    '19/06/2009
                    If Modo = 4 Then Exit Sub
                    '19/06/2009
                    
                    If txtAux(5).Text = "" Then Exit Sub
                    txtAux(9).Text = CargarPrecio(txtAux(7).Text, txtAux(5).Text)
                    txtAux(9).Text = Format(txtAux(9).Text, "##,##0.000")

                    Famia = DevuelveDesdeBD("codfamia", "sartic", "codartic", txtAux(7).Text, "N")
                    tipoF = DevuelveDesdeBD("tipfamia", "sfamia", "codfamia", Famia, "N")
                    If tipoF = "1" Then
                       If Modo = 3 Then BloquearTxt txtAux(8), True
                       If Modo = 3 Then BloquearTxt txtAux(10), False
                       PonerFoco txtAux(10)
                    Else
                       If Modo = 3 Then BloquearTxt txtAux(10), True
                       If Modo = 3 Then BloquearTxt txtAux(8), False
                       PonerFoco txtAux(8)
                    End If
                End If
            Else
                txtAux2(7).Text = ""
            End If
            
        Case 8 'CANTIDAD
            If PonerFormatoDecimal(txtAux(Index), 2) Then
                If Modo = 4 Then
                   CalcularImporteNue txtAux(8), txtAux(9), txtAux(10), 0  'Calcular es funcion
                Else
                    If tipoF <> "1" Then
                       txtAux(10).Text = CalcularImporte(txtAux(8).Text, txtAux(9).Text, txtAux(10).Text, tipoF) 'Calcular es funcion
                    End If
                End If
            End If
        
        Case 9 'PRECIO
            If PonerFormatoDecimal(txtAux(Index), 2) Then
                If Modo = 4 Then
                   CalcularImporteNue txtAux(8), txtAux(9), txtAux(10), 1
                
                Else
                    If tipoF = "1" Then
                       txtAux(8).Text = CalcularImporte(txtAux(8).Text, txtAux(9).Text, txtAux(10).Text, tipoF)
                       PonerFoco txtAux(11)
                    Else
                       txtAux(10).Text = CalcularImporte(txtAux(8).Text, txtAux(9).Text, txtAux(10).Text, tipoF)
                       PonerFoco txtAux(11)
                    End If
                End If
            End If
        
        '[Monica]24/08/2015
        Case 15 ' descuento alvic
            PonerFormatoDecimal txtAux(Index), 2
        
        Case 10 'IMPORTE
            If PonerFormatoDecimal(txtAux(Index), 3) Then
                If Modo = 4 Then
                   CalcularImporteNue txtAux(8), txtAux(9), txtAux(10), 2
                Else
                    If tipoF = "1" Then
                       txtAux(8).Text = CalcularImporte(txtAux(8).Text, txtAux(9).Text, txtAux(10).Text, tipoF) 'Calcular es funcion
                    End If
                End If
            End If
        
        Case 11 'trabajador
            If PonerFormatoEntero(txtAux(Index)) Then
                txtAux2(11).Text = PonerNombreDeCod(txtAux(Index), "straba", "nomtraba", "codtraba", "N")
                If txtAux2(11).Text = "" Then
                    cadMen = "No existe el Trabajador: " & txtAux(Index).Text & vbCrLf
                    cadMen = cadMen & "¿Desea crearlo?" & vbCrLf
                    If MsgBox(cadMen, vbQuestion + vbYesNo) = vbYes Then
                        Set frmTra = New frmManTraba
                        frmTra.DatosADevolverBusqueda = "0|1|"
                        frmTra.NuevoCodigo = txtAux(Index).Text
                        txtAux(Index).Text = ""
                        TerminaBloquear
                        frmTra.Show vbModal
                        Set frmTra = Nothing
                        If Modo = 4 Then BLOQUEADesdeFormulario2 Me, adodc1, 1
                    Else
                        txtAux(Index).Text = ""
                    End If
                    PonerFoco txtAux(Index)
                End If
            Else
                txtAux2(11).Text = ""
            End If

        Case 12 'forma de pago
            If PonerFormatoEntero(txtAux(Index)) Then
                txtAux2(12).Text = PonerNombreDeCod(txtAux(Index), "sforpa", "nomforpa", "codforpa", "N")
                If txtAux2(12).Text = "" Then
                    cadMen = "No existe la F.Pago: " & txtAux(Index).Text & vbCrLf
                    cadMen = cadMen & "¿Desea crearla?" & vbCrLf
                    If MsgBox(cadMen, vbQuestion + vbYesNo) = vbYes Then
                        Set frmFPa = New frmManFpago
                        frmFPa.DatosADevolverBusqueda = "0|1|"
                        frmFPa.NuevoCodigo = txtAux(Index).Text
                        txtAux(Index).Text = ""
                        TerminaBloquear
                        frmFPa.Show vbModal
                        Set frmFPa = Nothing
                        If Modo = 4 Then BLOQUEADesdeFormulario2 Me, adodc1, 1
                    Else
                        txtAux(Index).Text = ""
                    End If
                    PonerFoco txtAux(Index)
                End If
            Else
                txtAux2(12).Text = ""
            End If
        
        Case 14 ' kilometros
            PonerFormatoEntero txtAux(Index)
            
        Case 16 ' importe vale
            PonerFormatoDecimal txtAux(Index), 1
        
    
    End Select
End Sub

Private Function DatosOk() As Boolean
Dim b As Boolean
Dim SQL As String
Dim Fpag As String

    b = CompForm(Me)
    If Not b Then Exit Function
    
    If Modo = 3 Then   'Estamos insertando
        'comprobar si ya existe el campo de clave primaria
        If ExisteCP(txtAux(0)) Then b = False
        
    End If
    
    ' comprobamos que la tarjeta introducida esta asociada al socio
    SQL = ""
    SQL = DevuelveDesdeBDNew(cPTours, "starje", "numtarje", "codsocio", txtAux(5).Text, "N", , "numtarje", txtAux(6).Text, "N")
    If SQL = "" Then
        MsgBox "El número de Tarjeta introducida no corresponde al Socio. Revise.", vbExclamation
        PonerFoco txtAux(6)
        b = False
    End If
    
    '02/03/2007 añadido para dar aviso si cambian el socio de que la FP no se corresponde
    If Modo = 4 Then
        If txtAux(5).Text <> SocioAnt Then
            Fpag = ""
            Fpag = DevuelveDesdeBDNew(cPTours, "ssocio", "codforpa", "codsocio", txtAux(5).Text, "N")
            If CInt(Fpag) <> CInt(txtAux(12).Text) Then
                If MsgBox("La Forma de Pago no coincide con la del Cliente. ¿Desea modificarla?", vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
                    txtAux(12).Text = Fpag
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

Private Sub LimpiarCampos()
Dim I As Integer

    On Error Resume Next

    For I = 11 To 16
        txtAux(I).Text = ""
    Next I
    txtAux2(11).Text = ""
    txtAux2(12).Text = ""
    
    If Err.Number <> 0 Then Err.Clear
End Sub

Private Sub BotonImprimir()
    frmAlbTurno.CadB = CadB
    frmAlbTurno.Show vbModal
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

'' ### [DavidV] 26/04/2006: Activar/desactivar la rueda del ratón.
'Private Sub DataGrid1_GotFocus()
'  WheelHook DataGrid1
'End Sub
'
'Private Sub DataGrid1_Lostfocus()
'  WheelUnHook
'End Sub

Private Sub CalcularSumaPantalla()
Dim Rs As ADODB.Recordset
Dim SQL As String

  If Not adodc1.Recordset.EOF And CadB = "" Then CadB = "codclave > 0"
  If CadB <> "" Then
     SQL = "select sum(cantidad), sum(importel) FROM scaalb "
     SQL = SQL & " WHERE " & CadB
     Set Rs = New ADODB.Recordset ' Crear objeto
     Rs.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText ' abrir cursor
      If Not Rs.EOF Then
        SQL = "Cantidad: " & Format(Rs.Fields(0), "###,##0.000") & vbCrLf
        SQL = SQL & " Importe : " & Format(Rs.Fields(1), "####,##0.00")
        MsgBox "Totales Selección: " & vbCrLf & vbCrLf & SQL, vbInformation
      End If
     Rs.Close
     Set Rs = Nothing
    Else
        MsgBox "Haga primero una selección para ver Totales.", vbInformation
  End If
End Sub

Private Function CargarPrecio(Articulo As String, Socio As String) As String
Dim Tarifa As String
Dim precio As String

    Tarifa = ""

    If Socio <> "" Then
        Tarifa = DevuelveDesdeBD("codtarif", "ssocio", "codsocio", Socio, "N")
    End If

    precio = ""
    If Articulo <> "" Then
        If Tarifa <> "" Then
            precio = DevuelveDesdeBD("preventa", "starif", "codartic", Articulo, "N")
            If precio = "" Then
                ' en caso de que no haya precio de tarifa cogemos el PVP del articulo
                precio = DevuelveDesdeBD("preventa", "sartic", "codartic", Articulo, "N")
                If precio = "" Then precio = "0"
            End If
            CargarPrecio = precio
        End If
    End If

End Function

' ### [Monica] 10/10/2006
Private Sub MandaBusquedaTarjetas(CadB As String)
    Dim Cad As String
    Dim nReg As Long

    ' si hay mas de un registro llamamos al formulario
    Cad = "select count(*) from starje where " & CadB
    nReg = TotalRegistros(Cad)
    Select Case nReg
    Case 0
        MsgBox "Este cliente no tiene tarjeta asociada", vbExclamation
    Case 1
        txtAux(6).Text = DevuelveDesdeBD("numtarje", "starje", "codsocio", txtAux(5).Text, "N")
        txtAux(6).Text = Format(txtAux(6).Text, "00000000")
        Exit Sub
    Case Else
        'Cridem al form
        ' **************** arreglar-ho per a vore lo que es desije ****************
        ' NOTA: el total d'amples de ParaGrid, ha de sumar 100
        Cad = ""
        Cad = Cad & ParaGrid(txtAux(6), 19, "Tarjeta")
        Cad = Cad & "Titular|nomtarje|T||34·"
        Cad = Cad & "Tipo|CASE tiptarje WHEN 0 THEN ""Normal"" WHEN 1 THEN ""Bonificado"" WHEN 2 THEN ""Profesional"" END as tiptarje|T||10·"
        Cad = Cad & "Banco|codbanco|T||10·"
        Cad = Cad & "Sucur.|codsucur|T||10·"
        Cad = Cad & "DC|digcontr|T||4·"
        Cad = Cad & "Cuenta|cuentaba|T||15·"
        
    
    
        If Cad <> "" Then
            Screen.MousePointer = vbHourglass
            Set frmB = New frmBuscaGrid
            frmB.vCampos = Cad
            frmB.vTabla = "starje"
            frmB.vSQL = CadB
            HaDevueltoDatos = False
            frmB.vDevuelve = "0|" '*** els camps que volen que torne ***
            frmB.vTitulo = "Tarjetas" ' ***** repasa açò: títol de BuscaGrid *****
            frmB.vSelElem = 1
    
            frmB.Show vbModal
            Set frmB = Nothing
            'Si ha posat valors i tenim que es formulari de búsqueda llavors
            'tindrem que tancar el form llançant l'event
            If HaDevueltoDatos Then
                If (Not adodc1.Recordset.EOF) And DatosADevolverBusqueda <> "" Then _
                    cmdRegresar_Click
            Else   'de ha retornat datos, es a decir NO ha retornat datos
                PonerFoco txtAux(6)
            End If
        End If
   End Select
End Sub

' ### [Monica] 16/12/2010
Private Sub MandaBusquedaMatriculas(CadB As String)
    Dim Cad As String
    Dim nReg As Long

    ' si hay mas de un registro llamamos al formulario
    Cad = "select count(*) from smatri where " & CadB
    nReg = TotalRegistros(Cad)
    Select Case nReg
    Case 0
        MsgBox "Este cliente no tiene matrícula asociada", vbExclamation
    Case 1
        txtAux(13).Text = DevuelveDesdeBD("matricul", "smatri", "codsocio", txtAux(5).Text, "N")
        Exit Sub
    Case Else
        'Cridem al form
        ' **************** arreglar-ho per a vore lo que es desije ****************
        ' NOTA: el total d'amples de ParaGrid, ha de sumar 100
        Cad = ""
        Cad = Cad & ParaGrid(txtAux(13), 15, "Matrícula")
        Cad = Cad & "Observaciones|observac|T||85·"
    
        If Cad <> "" Then
            Screen.MousePointer = vbHourglass
            Set frmB1 = New frmBuscaGrid
            frmB1.vCampos = Cad
            frmB1.vTabla = "smatri"
            frmB1.vSQL = CadB
            HaDevueltoDatos = False
            frmB1.vDevuelve = "0|" '*** els camps que volen que torne ***
            frmB1.vTitulo = "Matrículas del Cliente" ' ***** repasa açò: títol de BuscaGrid *****
            frmB1.vSelElem = 0
    
            frmB1.Show vbModal
            Set frmB1 = Nothing
            'Si ha posat valors i tenim que es formulari de búsqueda llavors
            'tindrem que tancar el form llançant l'event
            If HaDevueltoDatos Then
                If (Not adodc1.Recordset.EOF) And DatosADevolverBusqueda <> "" Then _
                    cmdRegresar_Click
            Else   'de ha retornat datos, es a decir NO ha retornat datos
'                PonerFoco txtAux(6)
            End If
        End If
   End Select
End Sub




Private Function InsertarLinea(NumLinea As String) As Boolean
'Inserta un registro en la tabla de lineas de Albaranes: slialb
Dim SQL As String
Dim vWhere As String
Dim b As Boolean
Dim vCStock As CStock
Dim DentroTRANS As Boolean
Dim db As BaseDatos
Dim sql2 As String

Dim Existe As Boolean



    InsertarLinea = False
    SQL = ""
    DentroTRANS = False
        
    
    If vParamAplic.Cooperativa = 4 Then
        Set db = New BaseDatos
        db.abrir vSesion.CadenaConexion, "root", "aritel"
        db.Tipo = "MYSQL"
        db.AbrirTrans

        Set vCont = New CContador
        If vCont.ConseguirContador(CodTipoMov, DentroTRANS, db) Then
            txtAux(0).Text = vCont.Contador
            txtAux(1).Text = vCont.Contador
        
            Do
                sql2 = DevuelveDesdeBDNew(cPTours, "scaalb", "codclave", "codclave", txtAux(0).Text, "N")
                If sql2 <> "" Then
                    'Ya existe el contador incrementarlo
                    Existe = True
                    vCont.ConseguirContador CodTipoMov, DentroTRANS, db
                    txtAux(0).Text = vCont.Contador
                    txtAux(1).Text = vCont.Contador
                Else
                    Existe = False
                End If
            Loop Until Not Existe
        Else
            Exit Function
        End If
    End If
        
    'Conseguir el siguiente numero de linea
    vWhere = "scaalb.codclave = " & DBSet(txtAux(0).Text, "N")
'    numlinea = SugerirCodigoSiguienteStr(NomTablaLineas, "numlinea", vWhere)
    
    Set vCStock = New CStock
    If Not InicializarCStock(vCStock, "S", NumLinea) Then Exit Function
    
    If DatosOkLineaEnv(vCStock) Then 'Lineas de factura
        SQL = "INSERT INTO scaalb "
        SQL = SQL & "(codclave,codsocio,numtarje,numalbar,fecalbar,horalbar,codturno,codartic,cantidad,preciove,importel,codforpa,matricul,codtraba,numfactu,numlinea,declaradogp, kilometros, dtoalvic, importevale) "
        SQL = SQL & "VALUES (" & DBSet(txtAux(0).Text, "N") & ", " & DBSet(txtAux(5).Text, "N") & ", " & DBSet(txtAux(6).Text, "N") & ","
        SQL = SQL & DBSet(txtAux(1).Text, "T") & ", " ' albaran
        SQL = SQL & DBSet(txtAux(2).Text, "F") & ", " ' fecha albaran
        SQL = SQL & "'" & Format(txtAux(2).Text, "yyyy-mm-dd") & " " & Format(txtAux(3).Text, "hh:mm:ss") & "'," ' hora (datetime)
        SQL = SQL & DBSet(txtAux(4).Text, "N") & ", " & DBSet(txtAux(7).Text, "N") & ", " 'codturno, codartic
        SQL = SQL & DBSet(txtAux(8).Text, "N") & "," ' cantidad
        SQL = SQL & DBSet(txtAux(9).Text, "N") & "," ' precio de venta
        SQL = SQL & DBSet(txtAux(10).Text, "N") & "," ' importe
        SQL = SQL & DBSet(txtAux(12).Text, "N") & "," ' forpa
        SQL = SQL & DBSet(txtAux(13).Text, "T") & "," ' matricula
        SQL = SQL & DBSet(txtAux(11).Text, "N") & "," ' trabajador
        SQL = SQL & "0,0,0, " ' numfactu, numlinea, declaradogp
        '[Monica]24/06/2013: kilometros
        SQL = SQL & DBSet(txtAux(14).Text, "N") & "," '& ")"
        '[Monica]24/08/2015: faltaba introducir el dtoalvic de regaixo
        SQL = SQL & DBSet(txtAux(15).Text, "N") & ","
        SQL = SQL & DBSet(txtAux(16).Text, "N") & ")"
        
    Else
        Exit Function
    End If
    
    On Error GoTo EInsertarLinea
    If SQL <> "" Then
        Conn.BeginTrans
        DentroTRANS = True
        
        'insertar la linea
        Conn.Execute SQL
        
        'si hay control de stock para el articulo actualizar en salmac e insertar en smoval
        'en actualizar stock comprobamos si el articulo tiene control de stock
        b = vCStock.ActualizarStock
    End If
    Set vCStock = Nothing
    
    If b Then
        Conn.CommitTrans
        InsertarLinea = True
    Else
        Conn.RollbackTrans
        vCont.DevolverContador CodTipoMov, vCont.Contador, db
        InsertarLinea = False
    End If
    
    Set vCont = Nothing
    Set db = Nothing
    
    Exit Function

EInsertarLinea:
    InsertarLinea = False
    If DentroTRANS Then Conn.RollbackTrans
    If vParamAplic.Cooperativa = 4 Then vCont.DevolverContador CodTipoMov, vCont.Contador, db
    Set vCont = Nothing
    Set db = Nothing
    MuestraError Err.Number, "Insertar Lineas Albaranes" & vbCrLf & Err.Description
End Function


Private Function ModificarLinea() As Boolean
'Modifica registre en les taules de Llínies
Dim nomFrame As String
Dim V As Integer
Dim Cad As String
Dim SQL As String
Dim vCStock As CStock
Dim b As Boolean
Dim Mens As String
    
    On Error GoTo EModificarLinea

    ModificarLinea = False
    SQL = ""

        
    Set vCStock = New CStock
    If Not InicializarCStock(vCStock, "S") Then Exit Function
    
    If DatosOkLineaEnv(vCStock) Then
        '#### LAURA 15/11/2006
        Conn.BeginTrans
        
'        Set vCStock = New CStock
        'iniciamos la clase con los valores anteriores para deshacer lo q insertamos antes
        b = InicializarCStock(vCStock, "E")
        If b Then
            b = vCStock.DevolverStock 'eliminamos de smoval y devolvemos stock valores anteriores
            'ahora leemos los valores nuevos
            If b Then b = InicializarCStock(vCStock, "S")
            'insertamos en smoval y actualizamos stock a los valores nuevos
            vCStock.cantidad = CSng(ComprobarCero(txtAux(8).Text))
            If b Then b = vCStock.ActualizarStock
    
            '[Monica]12/04/2017: quitamos el tag del colectivo
            txtAux(17).Tag = ""
            
            'actualizar la linea de Albaran
            If b Then b = ModificaDesdeFormulario2(Me)
'                Sql = "UPDATE slialb Set codalmac = " & txtAux(4).Text & ", codartic=" & DBSet(txtAux(5).Text, "T") & ", "
'                Sql = Sql & "ampliaci=" & DBSet(Text2(16).Text, "T") & ", "
'                Sql = Sql & "cantidad= " & DBSet(txtAux(6).Text, "N") & ", "
'                Sql = Sql & "precioar= " & DBSet(txtAux(7).Text, "N") & ", " 'precio
'                Sql = Sql & "dtolinea= " & DBSet(txtAux(8).Text, "N") & ", "
'                Sql = Sql & "importel= " & DBSet(txtAux(9).Text, "N") & ", " 'Importe
'                Sql = Sql & "codigiva= " & DBSet(txtAux(10).Text, "N") & " " 'codigo de iva
'                Sql = Sql & Replace(ObtenerWhereCP(True), NombreTabla, "slialb") & " AND numlinea=" & AdoAux(1).Recordset!numlinea
'                Conn.Execute Sql
'            End If

            '[Monica]12/04/2017: lo volvemos a poner para la busqueda
            txtAux(17).Tag = "Colectivo|N|S|||ssocio|codcoope|000||"



            '[Monica]24/06/2013: modificamos la fechahora la dejaba en 1899-12-30
            If b Then
                SQL = "update scaalb set horalbar = '" & Format(txtAux(2).Text, "yyyy-mm-dd") & " " & Format(txtAux(3).Text, "hh:mm:ss") & "' "
                SQL = SQL & " where codclave = " & DBSet(txtAux(0).Text, "N")
                
                Conn.Execute SQL
            End If

        End If
    End If
    Set vCStock = Nothing
        
        
EModificarLinea:
    If Err.Number <> 0 Or Not b Then
        MuestraError Err.Number, "Modificar Lineas Factura" & vbCrLf & Err.Description & vbCrLf & Mens
        b = False
    End If
    
    If b Then
        Conn.CommitTrans
        ModificarLinea = True
    Else
        Conn.RollbackTrans
        ModificarLinea = False
    End If
End Function


Private Function InicializarCStock(ByRef vCStock As CStock, TipoM As String, Optional NumLinea As String) As Boolean
    On Error Resume Next

    vCStock.tipoMov = TipoM
    vCStock.DetaMov = "ALV" 'CodTipoMov 'Text1(6).Text
    vCStock.Trabajador = CLng(txtAux(5).Text) 'guardamos el cliente de la factura
    vCStock.Documento = txtAux(1).Text 'Nº albaran
    vCStock.Fechamov = txtAux(2).Text 'Fecha del albaran
    
    '1=Insertar, 2=Modificar
    If Modo = 3 Or (Modo = 4 And TipoM = "S") Then
        vCStock.codArtic = txtAux(7).Text
        vCStock.codAlmac = 1
        If Modo = 3 Then '1=Insertar
            vCStock.cantidad = CSng(ComprobarCero(txtAux(8).Text))
        Else '2=Modificar(Debe haber en stock la diferencia)
            If adodc1.Recordset!codArtic = txtAux(7).Text Then
                vCStock.cantidad = CSng(ComprobarCero(txtAux(8).Text)) - adodc1.Recordset!cantidad
            Else
                vCStock.cantidad = CSng(ComprobarCero(txtAux(8).Text))
            End If
        End If
        vCStock.Importe = CCur(ComprobarCero(txtAux(10).Text))
    Else
        vCStock.codArtic = adodc1.Recordset!codArtic
        vCStock.codAlmac = 1
        vCStock.cantidad = CSng(adodc1.Recordset!cantidad)
        vCStock.Importe = CCur(adodc1.Recordset!importel)
    End If
    If Modo = 3 Then
        vCStock.LineaDocu = 0 'CInt(ComprobarCero(numlinea))
    Else
        vCStock.LineaDocu = CInt(DevuelveValor("select numlinea from scaalb where codclave = " & DBSet(adodc1.Recordset!Codclave, "N")))
    End If
    If Err.Number <> 0 Then
        MsgBox "No se han podido inicializar la clase para actualizar Stock", vbExclamation
        InicializarCStock = False
    Else
        InicializarCStock = True
    End If
End Function

Private Function DatosOkLineaEnv(ByRef vCStock As CStock) As Boolean
Dim b As Boolean
Dim I As Byte
    
    On Error GoTo EDatosOkLineaEnv

    DatosOkLineaEnv = False
    b = True

    'Comprobar si se mueve mover stock (hay stock, o si no hay pero no control de stock)
    If vCStock.MueveStock Then
        b = vCStock.MoverStock
    End If
    DatosOkLineaEnv = b
    
EDatosOkLineaEnv:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Function

Private Function EliminarLinea() As Boolean
Dim SQL As String, Letra As String
Dim b As Boolean
Dim Mens As String
Dim vCStock As CStock

    On Error GoTo FinEliminar

    b = False
    If adodc1.Recordset.EOF Then Exit Function
        
    Conn.BeginTrans
        
    Mens = ""
    
        
    SQL = "Delete from scaalb where codclave=" & adodc1.Recordset!Codclave

     ' borramos el movimiento y aumentamos el stock
    Set vCStock = New CStock
    
    
    vCStock.tipoMov = "E"
    vCStock.DetaMov = "ALV" 'CodTipoMov 'Text1(6).Text
    vCStock.Trabajador = CLng(adodc1.Recordset!codsocio) 'guardamos el cliente de la factura
    vCStock.Documento = adodc1.Recordset!numalbar 'Nº albaran
    vCStock.Fechamov = adodc1.Recordset!fecAlbar 'Fecha del albaran
    
    vCStock.codArtic = adodc1.Recordset!codArtic
    vCStock.codAlmac = 1
    vCStock.cantidad = CSng(adodc1.Recordset!cantidad)
    vCStock.Importe = CCur(adodc1.Recordset!importel)
    vCStock.LineaDocu = CInt(DevuelveValor("select numlinea from scaalb where codclave = " & adodc1.Recordset!Codclave))
    
     'en actualizar stock comprobamos si el articulo tiene control de stock
     b = vCStock.DevolverStock
     Set vCStock = Nothing

     If vParamAplic.Cooperativa = 4 Then
        Set vCont = New CContador
        vCont.DevolverContador CodTipoMov, Val(txtAux(0).Text)
        Set vCont = Nothing
     End If


     Conn.Execute SQL
    
    
FinEliminar:
    If Err.Number <> 0 Or Not b Then
        MuestraError Err.Number, "Eliminar Albarán ", Err.Description & " " & Mens
        b = False
    End If
    If Not b Then
        Conn.RollbackTrans
        EliminarLinea = False
    Else
        Conn.CommitTrans
        EliminarLinea = True
    End If
End Function



'***************************************
Private Sub ActivaTicket()
    ImpresoraDefecto = Printer.DeviceName
    XPDefaultPrinter vParamAplic.ImpresoraTickets
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
    Dim R As Long
    ' Get the printer information for the currently selected
    ' printer in the list. The information is taken from the
    ' WIN.INI file.
    Buffer = Space(1024)
    R = GetProfileString("PrinterPorts", PrinterName, "", _
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
    Dim R As Long
    Dim L As Long
    DeviceLine = PrinterName & "," & DriverName & "," & PrinterPort
    ' Store the new printer information in the [WINDOWS] section of
    ' the WIN.INI file for the DEVICE= item
    R = WriteProfileString("windows", "Device", DeviceLine)
    ' Cause all applications to reload the INI file:
    L = SendMessage(HWND_BROADCAST, WM_WININICHANGE, 0, "windows")
End Sub
'------------------ Fin de los procesos relacionados con el cambio de impresora ----

