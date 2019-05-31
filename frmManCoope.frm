VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmManCoope 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Colectivos"
   ClientHeight    =   6285
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10890
   Icon            =   "frmManCoope.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6285
   ScaleWidth      =   10890
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   5115
      Left            =   0
      TabIndex        =   18
      Top             =   450
      Width           =   10815
      _ExtentX        =   19076
      _ExtentY        =   9022
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Datos Básicos"
      TabPicture(0)   =   "frmManCoope.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame3"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame2(2)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame2(0)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Frame2(1)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "Plantillas para Tarjetas"
      TabPicture(1)   =   "frmManCoope.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "text1(13)"
      Tab(1).Control(1)=   "text1(12)"
      Tab(1).Control(2)=   "text1(11)"
      Tab(1).Control(3)=   "text1(10)"
      Tab(1).Control(4)=   "Label1(14)"
      Tab(1).Control(5)=   "Label1(11)"
      Tab(1).Control(6)=   "Label1(10)"
      Tab(1).Control(7)=   "Label1(9)"
      Tab(1).ControlCount=   8
      Begin VB.TextBox text1 
         Height          =   285
         Index           =   13
         Left            =   -74340
         MaxLength       =   250
         TabIndex        =   37
         Tag             =   "Fichero tarjetas GP|T|S|||scoope|fichrpt04|||"
         Top             =   3450
         Width           =   9495
      End
      Begin VB.TextBox text1 
         Height          =   285
         Index           =   12
         Left            =   -74340
         MaxLength       =   250
         TabIndex        =   36
         Tag             =   "Fichero BM|T|S|||scoope|fichrpt03|||"
         Top             =   2700
         Width           =   9495
      End
      Begin VB.TextBox text1 
         Height          =   285
         Index           =   11
         Left            =   -74340
         MaxLength       =   250
         TabIndex        =   35
         Tag             =   "Fichero anverso y BM|T|S|||scoope|fichrpt02|||"
         Top             =   1890
         Width           =   9495
      End
      Begin VB.TextBox text1 
         Height          =   285
         Index           =   10
         Left            =   -74340
         MaxLength       =   250
         TabIndex        =   34
         Tag             =   "Fichero anverso|T|S|||scoope|fichrpt01|||"
         Top             =   1170
         Width           =   9495
      End
      Begin VB.Frame Frame2 
         Caption         =   "Datos básicos"
         ForeColor       =   &H00972E0B&
         Height          =   2325
         Index           =   1
         Left            =   120
         TabIndex        =   28
         Top             =   1395
         Width           =   5805
         Begin VB.TextBox text1 
            Height          =   285
            Index           =   2
            Left            =   960
            MaxLength       =   9
            TabIndex        =   2
            Tag             =   "NIF|T|S|||scoope|nifcoope|||"
            Top             =   360
            Width           =   1335
         End
         Begin VB.TextBox text1 
            Height          =   285
            Index           =   3
            Left            =   960
            MaxLength       =   40
            TabIndex        =   3
            Tag             =   "Domicilio|T|S|||scoope|domcoope|||"
            Top             =   720
            Width           =   4575
         End
         Begin VB.TextBox text1 
            Height          =   285
            Index           =   5
            Left            =   960
            MaxLength       =   35
            TabIndex        =   5
            Tag             =   "Población|T|S|||scoope|pobcoope|||"
            Top             =   1440
            Width           =   4575
         End
         Begin VB.TextBox text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   4
            Left            =   960
            MaxLength       =   6
            TabIndex        =   4
            Tag             =   "Código Postal|T|S|||scoope|codposta|||"
            Top             =   1080
            Width           =   735
         End
         Begin VB.TextBox text1 
            Height          =   285
            Index           =   6
            Left            =   960
            MaxLength       =   30
            TabIndex        =   6
            Tag             =   "Provincia|T|S|||scoope|procoope|||"
            Top             =   1800
            Width           =   4575
         End
         Begin VB.Label Label1 
            Caption         =   "N.I.F."
            Height          =   255
            Index           =   0
            Left            =   200
            TabIndex        =   33
            Top             =   360
            Width           =   495
         End
         Begin VB.Label Label1 
            Caption         =   "Domicilio"
            Height          =   255
            Index           =   2
            Left            =   195
            TabIndex        =   32
            Top             =   720
            Width           =   735
         End
         Begin VB.Label Label1 
            Caption         =   "Población"
            Height          =   255
            Index           =   4
            Left            =   195
            TabIndex        =   31
            Top             =   1440
            Width           =   735
         End
         Begin VB.Label Label1 
            Caption         =   "C.P."
            Height          =   255
            Index           =   7
            Left            =   195
            TabIndex        =   30
            Top             =   1080
            Width           =   495
         End
         Begin VB.Label Label1 
            Caption         =   "Provincia"
            Height          =   255
            Index           =   8
            Left            =   195
            TabIndex        =   29
            Top             =   1800
            Width           =   735
         End
      End
      Begin VB.Frame Frame2 
         Height          =   885
         Index           =   0
         Left            =   120
         TabIndex        =   25
         Top             =   450
         Width           =   10575
         Begin VB.TextBox text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   0
            Left            =   240
            MaxLength       =   3
            TabIndex        =   0
            Tag             =   "Código Colectivo|N|N|0|999|scoope|codcoope|000|S|"
            Top             =   400
            Width           =   735
         End
         Begin VB.TextBox text1 
            Height          =   285
            Index           =   1
            Left            =   1226
            MaxLength       =   40
            TabIndex        =   1
            Tag             =   "Nombre|T|N|||scoope|nomcoope|||"
            Top             =   400
            Width           =   4335
         End
         Begin VB.Label Label1 
            Caption         =   "Nombre "
            Height          =   255
            Index           =   1
            Left            =   1226
            TabIndex        =   27
            Top             =   195
            Width           =   735
         End
         Begin VB.Label Label1 
            Caption         =   "Cód."
            Height          =   255
            Index           =   3
            Left            =   240
            TabIndex        =   26
            Top             =   200
            Width           =   375
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Datos contacto"
         ForeColor       =   &H00972E0B&
         Height          =   2325
         Index           =   2
         Left            =   6120
         TabIndex        =   21
         Top             =   1395
         Width           =   4575
         Begin VB.TextBox text1 
            Height          =   285
            Index           =   7
            Left            =   960
            MaxLength       =   10
            TabIndex        =   7
            Tag             =   "Teléfono|T|S|||scoope|telcoope|||"
            Top             =   320
            Width           =   1335
         End
         Begin VB.TextBox text1 
            Height          =   285
            Index           =   8
            Left            =   960
            MaxLength       =   10
            TabIndex        =   8
            Tag             =   "Fax|T|S|||scoope|faxcoope|||"
            Top             =   680
            Width           =   1335
         End
         Begin VB.TextBox text1 
            Height          =   285
            Index           =   9
            Left            =   240
            MaxLength       =   40
            TabIndex        =   9
            Tag             =   "E-mail|T|S|||scoope|maicoope|||"
            Top             =   1320
            Width           =   4215
         End
         Begin VB.Label Label1 
            Caption         =   "Fax"
            Height          =   255
            Index           =   12
            Left            =   240
            TabIndex        =   24
            Top             =   680
            Width           =   735
         End
         Begin VB.Label Label1 
            Caption         =   "Teléfono"
            Height          =   255
            Index           =   13
            Left            =   240
            TabIndex        =   23
            Top             =   320
            Width           =   735
         End
         Begin VB.Label Label1 
            Caption         =   "E-mail"
            Height          =   255
            Index           =   16
            Left            =   240
            TabIndex        =   22
            Top             =   1120
            Width           =   735
         End
         Begin VB.Image imgMail 
            Height          =   240
            Index           =   0
            Left            =   960
            Top             =   1050
            Width           =   240
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Datos Facturación"
         ForeColor       =   &H00972E0B&
         Height          =   945
         Left            =   120
         TabIndex        =   19
         Top             =   3870
         Width           =   10545
         Begin VB.ComboBox Combo1 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   0
            ItemData        =   "frmManCoope.frx":0044
            Left            =   1800
            List            =   "frmManCoope.frx":0046
            Style           =   2  'Dropdown List
            TabIndex        =   10
            Tag             =   "Tipo Factura|N|N|0|4|scoope|tipfactu||N|"
            Top             =   330
            Width           =   1725
         End
         Begin VB.Label Label1 
            Caption         =   "Tipo Facturacion "
            Height          =   255
            Index           =   5
            Left            =   210
            TabIndex        =   20
            Top             =   360
            Width           =   1545
         End
      End
      Begin VB.Label Label1 
         Caption         =   "Fichero para impresión de tarjetas Gasóleo Profesional"
         Height          =   255
         Index           =   14
         Left            =   -74340
         TabIndex        =   41
         Top             =   3150
         Width           =   4395
      End
      Begin VB.Label Label1 
         Caption         =   "Fichero para impresión Banda Magnética sólo"
         Height          =   255
         Index           =   11
         Left            =   -74340
         TabIndex        =   40
         Top             =   2370
         Width           =   4395
      End
      Begin VB.Label Label1 
         Caption         =   "Fichero para impresión de Anverso y Banda Magnética"
         Height          =   255
         Index           =   10
         Left            =   -74340
         TabIndex        =   39
         Top             =   1590
         Width           =   4395
      End
      Begin VB.Label Label1 
         Caption         =   "Fichero para impresión de Anverso"
         Height          =   255
         Index           =   9
         Left            =   -74340
         TabIndex        =   38
         Top             =   870
         Width           =   4395
      End
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   8400
      TabIndex        =   11
      Top             =   5730
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   9660
      TabIndex        =   13
      Top             =   5730
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton cmdRegresar 
      Caption         =   "&Regresar"
      Height          =   375
      Left            =   9660
      TabIndex        =   15
      Top             =   5730
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Height          =   555
      Index           =   1
      Left            =   30
      TabIndex        =   12
      Top             =   5565
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
         TabIndex        =   14
         Top             =   240
         Width           =   2295
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
      TabIndex        =   16
      Top             =   0
      Width           =   10890
      _ExtentX        =   19209
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
      Begin VB.CheckBox chkVistaPrevia 
         Caption         =   "Vista previa"
         Height          =   195
         Left            =   8400
         TabIndex        =   17
         Top             =   120
         Value           =   2  'Grayed
         Width           =   1215
      End
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
Attribute VB_Name = "frmManCoope"
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

Private HaDevueltoDatos As Boolean
Private CadenaSelect As String
Private CadenaConsulta As String
Private CadB As String

Private WithEvents frmB As frmBuscaGrid
Attribute frmB.VB_VarHelpID = -1

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

Dim btnPrimero As Byte 'Variable que indica el nº del Botó PrimerRegistro en la Toolbar1
Dim indice As Byte 'Indice del text1 donde se ponen los datos devueltos desde otros Formularios de Mtos

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
    If DatosADevolverBusqueda <> "" Then
        cmdRegresar.visible = (Modo = 2)
    Else
        cmdRegresar.visible = False
    End If
    
    
    BloquearText1 Me, Modo
    BloquearCombo Me, Modo
    
    'Si es regresar
'    If DatosADevolverBusqueda <> "" Then cmdRegresar.visible = b
    
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
    
    LimpiarCampos 'Vacía los TextBox
    CadB = ""
    
    'Añadiremos el boton de aceptar y demas objetos para insertar
    PonerModo 3
    
     '******************** canviar taula i camp **************************
    If (DatosADevolverBusqueda <> "") And NuevoCodigo <> "" Then
        NumF = NuevoCodigo
    Else
        NumF = SugerirCodigoSiguienteStr("scoope", "codcoope")
    End If
    
    ' ******* Canviar el nom de la taula, el nom de la clau primaria, i el
    ' nom del camp que te la clau primaria si no es Text1(0) *************
    Text1(0).Text = NumF
    FormateaCampo Text1(0)
    
    'PosarDescripcions
    PonerFoco Text1(1)
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
            frmB.vTitulo = "Colectivos"
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
    PonerFoco Text1(1)
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
    'If EsCodigoCero(CStr(adodc1.Recordset.Fields(0).Value), FormatoCod) Then Exit Sub
    
    If Not SepuedeBorrar Then Exit Sub
    
    '*************** canviar els noms i el DELETE **********************************
    SQL = "¿Seguro que desea eliminar el Colectivo?"
    SQL = SQL & vbCrLf & "Código: " & Text1(0).Text
    SQL = SQL & vbCrLf & "Nombre: " & adodc1.Recordset.Fields(1)
    
    If MsgBox(SQL, vbQuestion + vbYesNo) = vbYes Then
        'Hay que eliminar
        Screen.MousePointer = vbHourglass
        NumRegElim = adodc1.Recordset.AbsolutePosition
        
        SQL = "Delete from " & NomTabla & " where codcoope=" & adodc1.Recordset!codcoope
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


Private Sub cmdAceptar_Click()

    Select Case Modo
         Case 1  'BUSQUEDA
            HacerBusqueda
    
        Case 3 'INSERTAR
            If DatosOk Then
                If InsertarDesdeForm(Me) Then
                    CadenaConsulta = "select * from " & NomTabla
                    CadenaConsulta = CadenaConsulta & " WHERE codcoope=" & Text1(0).Text
                    CadenaConsulta = CadenaConsulta & Ordenacion
                    Me.adodc1.RecordSource = CadenaConsulta '"Select * from " & NomTabla & Ordenacion
                    Me.adodc1.Refresh
                    PosicionarData
                End If
            End If
        
        Case 4 'MODIFICAR
            If DatosOk Then
                If ModificaDesdeFormulario(Me) Then
                    TerminaBloquear
                    PosicionarData
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

Private Sub CargaCombo()
    Combo1(0).Clear

    Combo1(0).AddItem "Tarjeta"
    Combo1(0).ItemData(Combo1(0).NewIndex) = 0
    Combo1(0).AddItem "Cliente"
    Combo1(0).ItemData(Combo1(0).NewIndex) = 1
    'VRS:2.0.2(1) añadida nueva opción
    Combo1(0).AddItem "Facturacion Ajena"
    Combo1(0).ItemData(Combo1(0).NewIndex) = 2
    Combo1(0).AddItem "Interna"
    Combo1(0).ItemData(Combo1(0).NewIndex) = 3
    
    Combo1(0).AddItem "Departamento"
    Combo1(0).ItemData(Combo1(0).NewIndex) = 4

'[Monica]25/09/2014: el tipo de contabilizacion pasa a estar en el socio en lugar de en el colectivo
'    Combo1(1).Clear
'
'    Combo1(1).AddItem "Cta.Contable Socio"
'    Combo1(1).ItemData(Combo1(1).NewIndex) = 0
'    Combo1(1).AddItem "Cta.Contable Cliente"
'    Combo1(1).ItemData(Combo1(1).NewIndex) = 1

End Sub

Private Sub Combo1_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub Form_Activate()
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
    
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

    'cargar IMAGE de mail
    Me.ImgMail(0).Picture = frmPpal.imgListImages16.ListImages(2).Picture
    CargaCombo

    Me.SSTab1.Tab = 0
    Me.SSTab1.TabVisible(1) = (vParamAplic.Cooperativa = 1)

    LimpiarCampos   'Limpia los campos TextBox
    
    'Vemos como esta guardado el valor del check
    chkVistaPrevia.Value = CheckValueLeer(Name)

    '****************** canviar la consulta *********************************+
    NomTabla = "scoope"
    Ordenacion = " ORDER BY codcoope"
    CadenaConsulta = "select * from " & NomTabla
    
    Me.adodc1.ConnectionString = Conn
    Me.adodc1.RecordSource = CadenaConsulta & " where codcoope=-1"
    Me.adodc1.Refresh
    
    CadB = ""

    If DatosADevolverBusqueda = "" Then
        PonerModo 0
    Else
        PonerModo 1
        Text1(0).BackColor = vbYellow 'codclien
    End If

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
    If adodc1.Recordset.EOF Then Exit Sub
    If adodc1.Recordset.RecordCount < 1 Then Exit Sub
    
    ' ### [Monica] 26/09/2006 dejamos modificar y eliminar el codigo 0
    'El registro de codigo 0 no se puede Modificar ni Eliminar
    'If EsCodigoCero(CStr(adodc1.Recordset.Fields(0).Value), FormatoCod) Then Exit Sub

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
    KEYpress KeyAscii
End Sub

Private Sub Text1_LostFocus(Index As Integer)
Dim Nuevo As Boolean

    If Not PerderFocoGnral(Text1(Index), Modo) Then Exit Sub
    
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub

    Select Case Index
        Case 0 'codigo trabajador
            PonerFormatoEntero Text1(0)
        
        Case 1 'NOMBRE
            Text1(Index).Text = UCase(Text1(Index).Text)
        
        Case 2 'NIF
            Text1(Index).Text = UCase(Text1(Index).Text)
            ValidarNIF Text1(Index).Text
            
        Case 13 ' ultima impresion de tarjeta de la solapa de tarjetas
            cmdAceptar.SetFocus
            
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
        Case 11 'Imprimir
                mnImprimir_Click
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
    
    '-- Esto permanece para saber donde estamos
    lblIndicador.Caption = PonerContRegistros(Me.adodc1)
End Sub

Private Function DatosOk() As Boolean
Dim b As Boolean

    b = CompForm(Me)
    If Not b Then Exit Function
    
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
Dim I As Integer

    On Error Resume Next

    limpiar Me
    For I = 0 To Me.Combo1.Count - 1
        Me.Combo1(I).ListIndex = -1
    Next I
    
    ' ****************************************************
    
    If Err.Number <> 0 Then Err.Clear
End Sub

Private Sub PosicionarData()
Dim Cad As String, Indicador As String

    Cad = "(codcoope=" & Text1(0).Text & ")"
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
        .cadTabla2 = "scoope"
        .Informe2 = "rManCoope.rpt"
        If CadB <> "" Then
            .cadRegSelec = SQL2SF(CadB)
        Else
            .cadRegSelec = ""
        End If
        .cadRegActua = POS2SF(adodc1, Me)
        .cadTodosReg = ""
        .OtrosParametros2 = "pEmpresa=" & DBSet(vEmpresa.nomEmpre, "T") & "|" & "'|pOrden={scoope.codcoope}|"
        .NumeroParametros2 = 1
        .MostrarTree2 = False
        .InfConta2 = False
        .ConSubInforme2 = False

        .Show vbModal
    End With
End Sub

