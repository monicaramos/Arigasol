VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmManClien 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Clientes"
   ClientHeight    =   7455
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11715
   Icon            =   "frmManClien.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7455
   ScaleWidth      =   11715
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Height          =   765
      Index           =   0
      Left            =   240
      TabIndex        =   59
      Top             =   480
      Width           =   11295
      Begin VB.TextBox text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   0
         Left            =   1440
         MaxLength       =   6
         TabIndex        =   0
         Tag             =   "Código de cliente|N|N|0|999999|ssocio|codsocio|000000|S|"
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox text1 
         Height          =   285
         Index           =   1
         Left            =   3480
         MaxLength       =   40
         TabIndex        =   1
         Tag             =   "Nombre|T|N|||ssocio|nomsocio|||"
         Top             =   240
         Width           =   4140
      End
      Begin VB.Label Label4 
         Caption         =   "Nombre "
         Height          =   255
         Left            =   2640
         TabIndex        =   61
         Top             =   240
         Width           =   660
      End
      Begin VB.Label Label1 
         Caption         =   "Código Cliente"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   60
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Height          =   555
      Index           =   0
      Left            =   240
      TabIndex        =   56
      Top             =   6840
      Width           =   2865
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
         Height          =   240
         Left            =   120
         TabIndex        =   57
         Top             =   180
         Width           =   2655
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   10410
      TabIndex        =   53
      Top             =   6960
      Width           =   1035
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   9120
      TabIndex        =   52
      Top             =   6960
      Width           =   1035
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5460
      Left            =   240
      TabIndex        =   58
      Top             =   1320
      Width           =   11295
      _ExtentX        =   19923
      _ExtentY        =   9631
      _Version        =   393216
      Style           =   1
      Tabs            =   5
      TabsPerRow      =   8
      TabHeight       =   520
      ShowFocusRect   =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Datos básicos"
      TabPicture(0)   =   "frmManClien.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label5"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label6(0)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label1(26)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label7"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label18"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label19"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "imgBuscar(0)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "imgBuscar(1)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label29"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "imgZoom(0)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "imgBuscar(2)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Label8"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Label1(1)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Label1(2)"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "text1(2)"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "text1(3)"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "text1(6)"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "text1(5)"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "text2(7)"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "text1(7)"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "text2(8)"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "text1(8)"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "text1(4)"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "FrameDatosAlta"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "FrameDatosContacto"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "text1(24)"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "text1(9)"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "text2(9)"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "text1(25)"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "text1(33)"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).ControlCount=   30
      TabCaption(1)   =   "Datos Tarjetas"
      TabPicture(1)   =   "frmManClien.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "FrameAux0"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Datos Matriculas"
      TabPicture(2)   =   "frmManClien.frx":0044
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "FrameAux1"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Márgenes"
      TabPicture(3)   =   "frmManClien.frx":0060
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "FrameAux2"
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "CRM"
      TabPicture(4)   =   "frmManClien.frx":007C
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "FrameA"
      Tab(4).Control(1)=   "cmdAccCRM(2)"
      Tab(4).Control(2)=   "cmdAccCRM(1)"
      Tab(4).Control(3)=   "cmdAccCRM(0)"
      Tab(4).Control(4)=   "lwCRM"
      Tab(4).Control(5)=   "Toolbar3"
      Tab(4).Control(6)=   "text1(28)"
      Tab(4).Control(7)=   "LabelCRM"
      Tab(4).ControlCount=   8
      Begin VB.TextBox text1 
         Height          =   285
         Index           =   33
         Left            =   4620
         MaxLength       =   10
         TabIndex        =   11
         Tag             =   "Codigo externo|T|S|||ssocio|codexterno|||"
         Top             =   3030
         Width           =   975
      End
      Begin VB.Frame FrameA 
         Caption         =   "Operaciones Aseguradas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00972E0B&
         Height          =   930
         Left            =   -74100
         TabIndex        =   113
         Top             =   4260
         Width           =   10005
         Begin VB.TextBox text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   30
            Left            =   8370
            MaxLength       =   40
            TabIndex        =   117
            Top             =   510
            Width           =   1425
         End
         Begin VB.TextBox text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   29
            Left            =   7200
            MaxLength       =   10
            TabIndex        =   116
            Top             =   510
            Width           =   1035
         End
         Begin VB.TextBox text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   27
            Left            =   4200
            MaxLength       =   10
            TabIndex        =   115
            Top             =   510
            Width           =   1035
         End
         Begin VB.TextBox text1 
            Height          =   285
            Index           =   26
            Left            =   1560
            MaxLength       =   40
            TabIndex        =   114
            Top             =   510
            Width           =   1455
         End
         Begin VB.Label Label27 
            Caption         =   "Fecha"
            Height          =   255
            Left            =   7200
            TabIndex        =   122
            Top             =   270
            Width           =   495
         End
         Begin VB.Label Label26 
            Caption         =   "Importe"
            Height          =   255
            Left            =   8400
            TabIndex        =   121
            Top             =   270
            Width           =   645
         End
         Begin VB.Label Label25 
            Caption         =   "Fecha Baja"
            Height          =   255
            Left            =   3210
            TabIndex        =   120
            Top             =   540
            Width           =   1005
         End
         Begin VB.Label Label24 
            Caption         =   "CONCEDIDO"
            Height          =   255
            Left            =   6060
            TabIndex        =   119
            Top             =   540
            Width           =   1125
         End
         Begin VB.Label Label20 
            Caption         =   "Nro.Póliza"
            Height          =   255
            Left            =   630
            TabIndex        =   118
            Top             =   510
            Width           =   885
         End
      End
      Begin VB.CommandButton cmdAccCRM 
         Height          =   375
         Index           =   2
         Left            =   -65040
         Picture         =   "frmManClien.frx":0098
         Style           =   1  'Graphical
         TabIndex        =   109
         ToolTipText     =   "Eliminar"
         Top             =   420
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.CommandButton cmdAccCRM 
         Height          =   375
         Index           =   1
         Left            =   -64440
         Picture         =   "frmManClien.frx":0A9A
         Style           =   1  'Graphical
         TabIndex        =   108
         ToolTipText     =   "Impresion CRM"
         Top             =   420
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.CommandButton cmdAccCRM 
         Height          =   375
         Index           =   0
         Left            =   -65520
         Picture         =   "frmManClien.frx":1024
         Style           =   1  'Graphical
         TabIndex        =   107
         ToolTipText     =   "Acciones CRM"
         Top             =   420
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Frame FrameAux2 
         BorderStyle     =   0  'None
         Height          =   4215
         Left            =   -74760
         TabIndex        =   97
         Top             =   480
         Width           =   10695
         Begin VB.TextBox txtAux 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   290
            Index           =   17
            Left            =   6300
            MaxLength       =   9
            TabIndex        =   101
            Tag             =   "Euros/litro|N|S|||smargen|euroslitro|#,##0.00000||"
            Text            =   "Euros/lit"
            Top             =   3600
            Visible         =   0   'False
            Width           =   1635
         End
         Begin VB.TextBox text2 
            BackColor       =   &H80000018&
            Enabled         =   0   'False
            Height          =   285
            Index           =   0
            Left            =   2340
            TabIndex        =   106
            Top             =   3600
            Visible         =   0   'False
            Width           =   2235
         End
         Begin VB.CommandButton btnBuscar 
            Appearance      =   0  'Flat
            Caption         =   "+"
            Height          =   290
            Index           =   0
            Left            =   2130
            MaskColor       =   &H00000000&
            TabIndex        =   105
            ToolTipText     =   "Buscar Artículo"
            Top             =   3600
            Visible         =   0   'False
            Width           =   195
         End
         Begin VB.TextBox txtAux 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   290
            Index           =   14
            Left            =   630
            MaxLength       =   2
            TabIndex        =   102
            Tag             =   "Numero linea|N|N|1|99|smargen|numlinea|00|S|"
            Text            =   "linea"
            Top             =   3600
            Visible         =   0   'False
            Width           =   555
         End
         Begin VB.TextBox txtAux 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   290
            Index           =   16
            Left            =   4590
            MaxLength       =   8
            TabIndex        =   100
            Tag             =   "Porcentaje|N|S|||smargen|margen|#,##0.00||"
            Text            =   "Porcenta"
            Top             =   3600
            Visible         =   0   'False
            Width           =   1635
         End
         Begin VB.TextBox txtAux 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   290
            Index           =   15
            Left            =   1320
            MaxLength       =   10
            TabIndex        =   99
            Tag             =   "Artículo|T|N|||smargen|codartic|000000||"
            Text            =   "codartic"
            Top             =   3600
            Visible         =   0   'False
            Width           =   795
         End
         Begin VB.TextBox txtAux 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   290
            Index           =   13
            Left            =   0
            MaxLength       =   6
            TabIndex        =   98
            Tag             =   "Código de cliente|N|N|1|999999|smargen|codsocio|000000|S|"
            Text            =   "codcli"
            Top             =   3600
            Visible         =   0   'False
            Width           =   555
         End
         Begin MSComctlLib.Toolbar ToolAux 
            Height          =   390
            Index           =   2
            Left            =   0
            TabIndex        =   103
            Top             =   0
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   688
            ButtonWidth     =   609
            ButtonHeight    =   582
            AllowCustomize  =   0   'False
            _Version        =   393216
            BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
               NumButtons      =   3
               BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Object.ToolTipText     =   "Nuevo"
                  Object.Tag             =   "2"
               EndProperty
               BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Object.ToolTipText     =   "Modificar"
                  Object.Tag             =   "2"
               EndProperty
               BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Object.ToolTipText     =   "Eliminar"
                  Object.Tag             =   "2"
               EndProperty
            EndProperty
         End
         Begin MSDataGridLib.DataGrid DataGridAux 
            Bindings        =   "frmManClien.frx":1A26
            Height          =   3645
            Index           =   2
            Left            =   0
            TabIndex        =   104
            Top             =   480
            Width           =   9645
            _ExtentX        =   17013
            _ExtentY        =   6429
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
         Begin MSAdodcLib.Adodc AdoAux 
            Height          =   375
            Index           =   2
            Left            =   4440
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
            Caption         =   "AdoAux(1)"
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
      End
      Begin VB.TextBox text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   25
         Left            =   1440
         MaxLength       =   6
         TabIndex        =   10
         Tag             =   "Nro Socio Cooperativa|N|S|||ssocio|nrosocio|000000||"
         Top             =   3015
         Width           =   735
      End
      Begin VB.Frame FrameAux1 
         BorderStyle     =   0  'None
         Height          =   4215
         Left            =   -74760
         TabIndex        =   92
         Top             =   480
         Width           =   10695
         Begin VB.TextBox txtAux 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   290
            Index           =   8
            Left            =   0
            MaxLength       =   6
            TabIndex        =   37
            Tag             =   "Código de cliente|N|N|1|999999|smatri|codsocio|000000|S|"
            Text            =   "codcli"
            Top             =   3600
            Visible         =   0   'False
            Width           =   555
         End
         Begin VB.TextBox txtAux 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   290
            Index           =   10
            Left            =   1320
            MaxLength       =   10
            TabIndex        =   39
            Tag             =   "Matricula|T|N|||smatri|matricul|||"
            Text            =   "matric"
            Top             =   3600
            Visible         =   0   'False
            Width           =   795
         End
         Begin VB.TextBox txtAux 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   290
            Index           =   11
            Left            =   2280
            MaxLength       =   30
            TabIndex        =   40
            Tag             =   "Observaciones|T|S|||smatri|observac|||"
            Text            =   "observac"
            Top             =   3600
            Visible         =   0   'False
            Width           =   1635
         End
         Begin VB.TextBox txtAux 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   290
            Index           =   9
            Left            =   600
            MaxLength       =   2
            TabIndex        =   38
            Tag             =   "Numero linea|N|N|1|99|smatri|numlinea|00|S|"
            Text            =   "linea"
            Top             =   3600
            Visible         =   0   'False
            Width           =   555
         End
         Begin MSComctlLib.Toolbar ToolAux 
            Height          =   390
            Index           =   1
            Left            =   0
            TabIndex        =   93
            Top             =   0
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   688
            ButtonWidth     =   609
            ButtonHeight    =   582
            AllowCustomize  =   0   'False
            _Version        =   393216
            BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
               NumButtons      =   3
               BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Object.ToolTipText     =   "Nuevo"
                  Object.Tag             =   "2"
               EndProperty
               BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Object.ToolTipText     =   "Modificar"
                  Object.Tag             =   "2"
               EndProperty
               BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Object.ToolTipText     =   "Eliminar"
                  Object.Tag             =   "2"
               EndProperty
            EndProperty
         End
         Begin MSDataGridLib.DataGrid DataGridAux 
            Bindings        =   "frmManClien.frx":1A3E
            Height          =   3645
            Index           =   1
            Left            =   0
            TabIndex        =   94
            Top             =   480
            Width           =   10515
            _ExtentX        =   18547
            _ExtentY        =   6429
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
         Begin MSAdodcLib.Adodc AdoAux 
            Height          =   375
            Index           =   1
            Left            =   4440
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
            Caption         =   "AdoAux(1)"
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
      End
      Begin VB.Frame FrameAux0 
         BorderStyle     =   0  'None
         Height          =   4395
         Left            =   -74880
         TabIndex        =   89
         Top             =   480
         Width           =   11055
         Begin VB.TextBox txtAux 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   290
            Index           =   18
            Left            =   4200
            MaxLength       =   4
            TabIndex        =   46
            Tag             =   "IBAN|T|S|||starje|iban|||"
            Text            =   "iban"
            Top             =   3720
            Visible         =   0   'False
            Width           =   435
         End
         Begin VB.TextBox txtAux 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   290
            Index           =   12
            Left            =   6840
            MaxLength       =   10
            TabIndex        =   51
            Tag             =   "Matricula|T|S|||starje|matricul|||"
            Text            =   "matricul"
            Top             =   3720
            Visible         =   0   'False
            Width           =   555
         End
         Begin VB.TextBox txtAux 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   290
            Index           =   3
            Left            =   1920
            MaxLength       =   40
            TabIndex        =   45
            Tag             =   "Titular|T|S|||starje|nomtarje|||"
            Text            =   "ta"
            Top             =   3720
            Visible         =   0   'False
            Width           =   2295
         End
         Begin VB.TextBox txtAux 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   290
            Index           =   0
            Left            =   -120
            MaxLength       =   6
            TabIndex        =   41
            Tag             =   "Código de cliente|N|N|1|999999|starje|codsocio|000000|S|"
            Text            =   "codc"
            Top             =   3720
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.TextBox txtAux 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   290
            Index           =   1
            Left            =   360
            MaxLength       =   2
            TabIndex        =   42
            Tag             =   "Número de línea|N|N|1|99|starje|numlinea|00|S|"
            Text            =   "nu"
            Top             =   3720
            Visible         =   0   'False
            Width           =   315
         End
         Begin VB.TextBox txtAux 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   290
            Index           =   4
            Left            =   4680
            MaxLength       =   4
            TabIndex        =   47
            Tag             =   "Banco|T|S|0|9999|starje|codbanco|0000||"
            Text            =   "banc"
            Top             =   3720
            Visible         =   0   'False
            Width           =   435
         End
         Begin VB.TextBox txtAux 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   290
            Index           =   5
            Left            =   5310
            MaxLength       =   4
            TabIndex        =   48
            Tag             =   "Oficina|T|S|0|9999|starje|codsucur|0000||"
            Text            =   "sucu"
            Top             =   3720
            Visible         =   0   'False
            Width           =   435
         End
         Begin VB.TextBox txtAux 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   290
            Index           =   6
            Left            =   5760
            MaxLength       =   2
            TabIndex        =   49
            Tag             =   "Dígito de Control|T|S|||starje|digcontr|00||"
            Text            =   "DC"
            Top             =   3720
            Visible         =   0   'False
            Width           =   315
         End
         Begin VB.TextBox txtAux 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   290
            Index           =   7
            Left            =   6120
            MaxLength       =   10
            TabIndex        =   50
            Tag             =   "Número de Cuenta|T|S|||starje|cuentaba|0000000000||"
            Text            =   "cuenta"
            Top             =   3720
            Visible         =   0   'False
            Width           =   555
         End
         Begin VB.TextBox txtAux 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   290
            Index           =   2
            Left            =   720
            MaxLength       =   8
            TabIndex        =   43
            Tag             =   "Tarjeta|N|N|1|99999999|starje|numtarje|00000000||"
            Text            =   "ta"
            Top             =   3720
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.ComboBox cmbAux 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   0
            ItemData        =   "frmManClien.frx":1A56
            Left            =   1200
            List            =   "frmManClien.frx":1A58
            Style           =   2  'Dropdown List
            TabIndex        =   44
            Tag             =   "Tipo Tarjeta|N|N|||starje|tiptarje|||"
            Top             =   3720
            Visible         =   0   'False
            Width           =   615
         End
         Begin MSComctlLib.Toolbar ToolAux 
            Height          =   390
            Index           =   0
            Left            =   0
            TabIndex        =   90
            Top             =   0
            Width           =   1560
            _ExtentX        =   2752
            _ExtentY        =   688
            ButtonWidth     =   609
            ButtonHeight    =   582
            AllowCustomize  =   0   'False
            _Version        =   393216
            BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
               NumButtons      =   5
               BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Object.ToolTipText     =   "Nuevo"
                  Object.Tag             =   "2"
               EndProperty
               BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Object.ToolTipText     =   "Modificar"
                  Object.Tag             =   "2"
               EndProperty
               BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Object.ToolTipText     =   "Eliminar"
                  Object.Tag             =   "2"
               EndProperty
               BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Style           =   3
               EndProperty
               BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Enabled         =   0   'False
                  Object.Visible         =   0   'False
                  Object.ToolTipText     =   "Tarjetas Libres"
                  Object.Tag             =   "2"
               EndProperty
            EndProperty
         End
         Begin MSAdodcLib.Adodc AdoAux 
            Height          =   375
            Index           =   0
            Left            =   3720
            Top             =   480
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
            Caption         =   "AdoAux(0)"
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
         Begin MSDataGridLib.DataGrid DataGridAux 
            Bindings        =   "frmManClien.frx":1A5A
            Height          =   3825
            Index           =   0
            Left            =   0
            TabIndex        =   91
            Top             =   480
            Width           =   10965
            _ExtentX        =   19341
            _ExtentY        =   6747
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
         Begin VB.Label Label23 
            Caption         =   "Impresión de Tarjetas"
            Height          =   255
            Left            =   8820
            TabIndex        =   125
            Top             =   60
            Width           =   1590
         End
         Begin VB.Image imgDoc 
            Height          =   405
            Index           =   1
            Left            =   10470
            ToolTipText     =   "Impresión de Tarjetas"
            Top             =   0
            Width           =   390
         End
      End
      Begin VB.TextBox text2 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   9
         Left            =   2700
         TabIndex        =   82
         Top             =   2640
         Width           =   2895
      End
      Begin VB.TextBox text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   9
         Left            =   1440
         MaxLength       =   10
         TabIndex        =   9
         Tag             =   "Cta.Contable|T|S|||ssocio|codmacta|||"
         Top             =   2640
         Width           =   1215
      End
      Begin VB.TextBox text1 
         Height          =   975
         Index           =   24
         Left            =   5880
         MaxLength       =   250
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   36
         Tag             =   "Observaciones|T|S|||ssocio|obssocio|||"
         Top             =   4320
         Width           =   5295
      End
      Begin VB.Frame FrameDatosContacto 
         Caption         =   "Datos Contacto"
         ForeColor       =   &H00972E0B&
         Height          =   1920
         Left            =   225
         TabIndex        =   71
         Top             =   3390
         Width           =   5415
         Begin VB.TextBox text1 
            Height          =   285
            Index           =   10
            Left            =   1200
            MaxLength       =   40
            TabIndex        =   12
            Tag             =   "Web|T|S|||ssocio|wwwsocio|||"
            Top             =   360
            Width           =   4095
         End
         Begin VB.TextBox text1 
            Height          =   285
            Index           =   11
            Left            =   1200
            MaxLength       =   10
            TabIndex        =   13
            Tag             =   "Teléfono|T|S|||ssocio|telsocio|||"
            Top             =   730
            Width           =   1455
         End
         Begin VB.TextBox text1 
            Height          =   285
            Index           =   12
            Left            =   3840
            MaxLength       =   10
            TabIndex        =   14
            Tag             =   "Móvil|T|S|||ssocio|movsocio|||"
            Top             =   730
            Width           =   1455
         End
         Begin VB.TextBox text1 
            Height          =   285
            Index           =   13
            Left            =   1200
            MaxLength       =   10
            TabIndex        =   15
            Tag             =   "Fax|T|S|||ssocio|faxsocio|||"
            Top             =   1100
            Width           =   1455
         End
         Begin VB.TextBox text1 
            Height          =   285
            Index           =   14
            Left            =   1200
            MaxLength       =   40
            TabIndex        =   16
            Tag             =   "E-mail|T|S|||ssocio|maisocio|||"
            Top             =   1440
            Width           =   4095
         End
         Begin VB.Image imgWeb 
            Height          =   240
            Index           =   0
            Left            =   735
            Picture         =   "frmManClien.frx":1A72
            Top             =   360
            Width           =   240
         End
         Begin VB.Label Label9 
            Caption         =   "Web"
            Height          =   255
            Left            =   240
            TabIndex        =   79
            Top             =   360
            Width           =   495
         End
         Begin VB.Image imgMail 
            Height          =   240
            Index           =   0
            Left            =   720
            Top             =   1470
            Width           =   240
         End
         Begin VB.Label Label10 
            Caption         =   "Teléfono"
            Height          =   255
            Left            =   240
            TabIndex        =   75
            Top             =   730
            Width           =   735
         End
         Begin VB.Label Label12 
            Caption         =   "Móvil"
            Height          =   255
            Left            =   3165
            TabIndex        =   74
            Top             =   730
            Width           =   495
         End
         Begin VB.Label Label14 
            Caption         =   "Fax"
            Height          =   255
            Left            =   240
            TabIndex        =   73
            Top             =   1100
            Width           =   735
         End
         Begin VB.Label Label16 
            Caption         =   "E-mail"
            Height          =   255
            Left            =   240
            TabIndex        =   72
            Top             =   1470
            Width           =   495
         End
      End
      Begin VB.Frame FrameDatosAlta 
         Caption         =   "Datos Facturación"
         ForeColor       =   &H00972E0B&
         Height          =   3585
         Left            =   5760
         TabIndex        =   68
         Top             =   360
         Width           =   5415
         Begin VB.CheckBox chkAux 
            Caption         =   "Contabilización sobre Cta Socio"
            Height          =   195
            Index           =   6
            Left            =   2550
            TabIndex        =   127
            Tag             =   "Tipo contab.|N|N|0|1|ssocio|tipconta||N|"
            Top             =   2070
            Width           =   2595
         End
         Begin VB.TextBox text1 
            Height          =   285
            Index           =   32
            Left            =   1290
            MaxLength       =   4
            TabIndex        =   20
            Tag             =   "IBAN|T|S|||ssocio|iban|||"
            Top             =   948
            Width           =   570
         End
         Begin VB.CheckBox chkAux 
            Caption         =   "Cliente de Varios"
            Height          =   195
            Index           =   5
            Left            =   2550
            TabIndex        =   33
            Tag             =   "Es de Varios|N|N|0|1|ssocio|esdevarios|||"
            Top             =   2820
            Width           =   2295
         End
         Begin VB.TextBox text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   31
            Left            =   1290
            MaxLength       =   2
            TabIndex        =   29
            Tag             =   "Dia Pago|N|S|1|31|ssocio|diapago|#0||"
            Top             =   2040
            Width           =   1080
         End
         Begin VB.CheckBox chkAux 
            Caption         =   "Bonificación Especial"
            Height          =   195
            Index           =   4
            Left            =   2550
            TabIndex        =   31
            Tag             =   "Bonif.Especial|N|N|0|1|ssocio|bonifesp|||"
            Top             =   2460
            Width           =   2295
         End
         Begin VB.CheckBox chkAux 
            Caption         =   "Bonificación Basica"
            Height          =   195
            Index           =   3
            Left            =   150
            TabIndex        =   30
            Tag             =   "Bonif.Basica|N|N|0|1|ssocio|bonifbas|||"
            Top             =   2460
            Width           =   2295
         End
         Begin VB.CheckBox chkAux 
            Caption         =   "Factura con FPago ficha"
            Height          =   195
            Index           =   2
            Left            =   150
            TabIndex        =   32
            Tag             =   "Factura con FP|N|N|0|1|ssocio|facturafp|||"
            Top             =   2820
            Width           =   2295
         End
         Begin VB.CheckBox chkAux 
            Caption         =   "Imprime Factura"
            Height          =   195
            Index           =   1
            Left            =   150
            TabIndex        =   34
            Tag             =   "Imprime Factura|N|N|0|1|ssocio|impfactu|||"
            Top             =   3180
            Width           =   2295
         End
         Begin VB.CheckBox chkAux 
            Caption         =   "Envio Factura por eMail"
            Height          =   195
            Index           =   0
            Left            =   2550
            TabIndex        =   35
            Tag             =   "Envio Factura eMail|N|N|0|1|ssocio|envfactemail|||"
            Top             =   3180
            Width           =   2295
         End
         Begin VB.ComboBox Combo1 
            Height          =   315
            Index           =   1
            ItemData        =   "frmManClien.frx":1FFC
            Left            =   3720
            List            =   "frmManClien.frx":2006
            Style           =   2  'Dropdown List
            TabIndex        =   28
            Tag             =   "Grupo Est.Artículos|N|N|0|1|ssocio|grupoestartic|||"
            Top             =   1650
            Width           =   1305
         End
         Begin VB.TextBox text1 
            Height          =   285
            Index           =   21
            Left            =   3720
            MaxLength       =   10
            TabIndex        =   24
            Tag             =   "Cuenta|T|S|||ssocio|cuentaba|||"
            Top             =   948
            Width           =   1290
         End
         Begin VB.TextBox text1 
            Height          =   285
            Index           =   20
            Left            =   3180
            MaxLength       =   2
            TabIndex        =   23
            Tag             =   "D.C.|T|S|||ssocio|digcontr|||"
            Top             =   948
            Width           =   480
         End
         Begin VB.TextBox text1 
            Height          =   285
            Index           =   19
            Left            =   2520
            MaxLength       =   4
            TabIndex        =   22
            Tag             =   "Sucursal|N|S|||ssocio|codsucur|0000||"
            Top             =   948
            Width           =   570
         End
         Begin VB.TextBox text1 
            Height          =   285
            Index           =   18
            Left            =   1890
            MaxLength       =   4
            TabIndex        =   21
            Tag             =   "Banco|N|S|||ssocio|codbanco|0000||"
            Top             =   948
            Width           =   570
         End
         Begin VB.TextBox text1 
            Height          =   285
            Index           =   23
            Left            =   3720
            MaxLength       =   8
            TabIndex        =   26
            Tag             =   "Dto./Litro|N|N|0|9.9999|ssocio|dtolitro|0.0000||"
            Top             =   1320
            Width           =   1290
         End
         Begin VB.TextBox text1 
            Height          =   285
            Index           =   22
            Left            =   1290
            MaxLength       =   1
            TabIndex        =   25
            Tag             =   "Tarifa|N|N|0|9|ssocio|codtarif|||"
            Top             =   1302
            Width           =   360
         End
         Begin VB.TextBox text2 
            BackColor       =   &H80000018&
            Enabled         =   0   'False
            Height          =   285
            Index           =   17
            Left            =   1890
            TabIndex        =   84
            Top             =   600
            Width           =   3135
         End
         Begin VB.TextBox text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   17
            Left            =   1290
            MaxLength       =   4
            TabIndex        =   19
            Tag             =   "Código F.Pago|N|N|0|999|ssocio|codforpa|000||"
            Top             =   594
            Width           =   555
         End
         Begin VB.ComboBox Combo1 
            Height          =   315
            Index           =   0
            ItemData        =   "frmManClien.frx":201C
            Left            =   1290
            List            =   "frmManClien.frx":2026
            Style           =   2  'Dropdown List
            TabIndex        =   27
            Tag             =   "Tipo de cliente|N|N|0|2|ssocio|tipsocio|||"
            Top             =   1650
            Width           =   1095
         End
         Begin VB.TextBox text1 
            Height          =   285
            Index           =   16
            Left            =   3810
            MaxLength       =   10
            TabIndex        =   18
            Tag             =   "F.Baja|F|S|||ssocio|fechabaj|dd/mm/yyyy||"
            Top             =   240
            Width           =   1200
         End
         Begin VB.TextBox text1 
            Height          =   285
            Index           =   15
            Left            =   1290
            MaxLength       =   10
            TabIndex        =   17
            Tag             =   "F.Alta|F|N|||ssocio|fechaalt|dd/mm/yyyy||"
            Top             =   240
            Width           =   1200
         End
         Begin VB.Label Label22 
            Caption         =   "Dia de Pago"
            Height          =   255
            Left            =   120
            TabIndex        =   124
            Top             =   2055
            Width           =   1035
         End
         Begin VB.Label Label2 
            Caption         =   "Grupo Est.Art."
            Height          =   255
            Index           =   1
            Left            =   2550
            TabIndex        =   95
            Top             =   1710
            Width           =   1005
         End
         Begin VB.Image imgFec 
            Height          =   240
            Index           =   16
            Left            =   3510
            Picture         =   "frmManClien.frx":203C
            ToolTipText     =   "Buscar fecha"
            Top             =   240
            Width           =   240
         End
         Begin VB.Image imgFec 
            Height          =   240
            Index           =   15
            Left            =   990
            Picture         =   "frmManClien.frx":20C7
            ToolTipText     =   "Buscar fecha"
            Top             =   240
            Width           =   240
         End
         Begin VB.Label Label17 
            Caption         =   "IBAN Cliente"
            Height          =   255
            Left            =   120
            TabIndex        =   88
            Top             =   975
            Width           =   975
         End
         Begin VB.Label Label15 
            Caption         =   "Dto. x Litro"
            Height          =   255
            Left            =   2550
            TabIndex        =   87
            Top             =   1335
            Width           =   855
         End
         Begin VB.Label Label13 
            Caption         =   "Tarifa"
            Height          =   255
            Left            =   120
            TabIndex        =   86
            Top             =   1335
            Width           =   855
         End
         Begin VB.Label Label11 
            Caption         =   "F.Pago"
            Height          =   255
            Left            =   120
            TabIndex        =   85
            Top             =   600
            Width           =   615
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   3
            Left            =   990
            ToolTipText     =   "Buscar F.Pago"
            Top             =   600
            Width           =   240
         End
         Begin VB.Label Label2 
            Caption         =   "Tipo Cliente"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   81
            Top             =   1680
            Width           =   975
         End
         Begin VB.Label Label3 
            Caption         =   "Fecha Baja"
            Height          =   255
            Left            =   2580
            TabIndex        =   80
            Top             =   255
            Width           =   825
         End
         Begin VB.Label Label21 
            Caption         =   "Fecha Alta"
            Height          =   255
            Left            =   120
            TabIndex        =   70
            Top             =   255
            Width           =   855
         End
      End
      Begin VB.TextBox text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   4
         Left            =   1440
         MaxLength       =   6
         TabIndex        =   4
         Tag             =   "C.Postal|T|N|||ssocio|codposta|||"
         Top             =   1226
         Width           =   735
      End
      Begin VB.TextBox text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   8
         Left            =   1440
         MaxLength       =   3
         TabIndex        =   8
         Tag             =   "Situación|N|N|0|99|ssocio|codsitua|00||"
         Top             =   2280
         Width           =   495
      End
      Begin VB.TextBox text2 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   8
         Left            =   1980
         TabIndex        =   55
         Top             =   2285
         Width           =   3615
      End
      Begin VB.TextBox text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   7
         Left            =   1440
         MaxLength       =   4
         TabIndex        =   7
         Tag             =   "Colectivo|N|N|0|999|ssocio|codcoope|000||"
         Top             =   1932
         Width           =   495
      End
      Begin VB.TextBox text2 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   7
         Left            =   1980
         TabIndex        =   54
         Top             =   1932
         Width           =   3615
      End
      Begin VB.TextBox text1 
         Height          =   285
         Index           =   5
         Left            =   2220
         MaxLength       =   35
         TabIndex        =   5
         Tag             =   "Población|T|N|||ssocio|pobsocio|||"
         Top             =   1226
         Width           =   3375
      End
      Begin VB.TextBox text1 
         Height          =   285
         Index           =   6
         Left            =   1440
         MaxLength       =   35
         TabIndex        =   6
         Tag             =   "Provincia|T|N|||ssocio|prosocio|||"
         Top             =   1579
         Width           =   4155
      End
      Begin VB.TextBox text1 
         Height          =   285
         Index           =   3
         Left            =   1440
         MaxLength       =   40
         TabIndex        =   3
         Tag             =   "Domicilio|T|N|||ssocio|domsocio|||"
         Top             =   873
         Width           =   4155
      End
      Begin VB.TextBox text1 
         Height          =   285
         Index           =   2
         Left            =   1440
         MaxLength       =   9
         TabIndex        =   2
         Tag             =   "NIF / CIF|T|N|||ssocio|nifsocio|||"
         Top             =   520
         Width           =   1200
      End
      Begin MSComctlLib.ListView lwCRM 
         Height          =   3345
         Left            =   -74160
         TabIndex        =   110
         Top             =   780
         Width           =   10095
         _ExtentX        =   17806
         _ExtentY        =   5900
         View            =   3
         LabelEdit       =   1
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin MSComctlLib.Toolbar Toolbar3 
         Height          =   1710
         Left            =   -74820
         TabIndex        =   112
         Top             =   780
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   3016
         ButtonWidth     =   609
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         Appearance      =   1
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   13
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Object.Visible         =   0   'False
               Object.ToolTipText     =   "Acciones comerciales"
               Object.Tag             =   "0"
               Style           =   2
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Object.Visible         =   0   'False
               Style           =   3
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Object.Visible         =   0   'False
               Object.ToolTipText     =   "Llamadas"
               Object.Tag             =   "1"
               Style           =   2
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Object.Visible         =   0   'False
               Style           =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Object.Visible         =   0   'False
               Object.ToolTipText     =   "Correo electronico"
               Object.Tag             =   "2"
               Style           =   2
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Object.Visible         =   0   'False
               Style           =   3
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Cobros"
               Object.Tag             =   "3"
               Style           =   2
               Value           =   1
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Object.Visible         =   0   'False
               Object.ToolTipText     =   "Observaciones departamento"
               Object.Tag             =   "4"
               Style           =   2
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Reclamaciones"
               Object.Tag             =   "5"
               Style           =   2
            EndProperty
            BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Historial"
               Object.Tag             =   "6"
               Style           =   2
            EndProperty
         EndProperty
      End
      Begin VB.TextBox text1 
         Height          =   285
         Index           =   28
         Left            =   -70740
         MaxLength       =   40
         TabIndex        =   123
         Top             =   4830
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Código Externo"
         Height          =   255
         Index           =   2
         Left            =   3390
         TabIndex        =   126
         Top             =   3060
         Width           =   1320
      End
      Begin VB.Label LabelCRM 
         Caption         =   "Label2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   300
         Left            =   -74160
         TabIndex        =   111
         Top             =   420
         Width           =   5745
      End
      Begin VB.Label Label1 
         Caption         =   "Nro.Socio"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   96
         Top             =   3015
         Width           =   870
      End
      Begin VB.Label Label8 
         Caption         =   "Cta.Conta."
         Height          =   255
         Left            =   240
         TabIndex        =   83
         Top             =   2640
         Width           =   735
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   2
         Left            =   1080
         ToolTipText     =   "Buscar Cta.Contable"
         Top             =   2640
         Width           =   240
      End
      Begin VB.Image imgZoom 
         Height          =   240
         Index           =   0
         Left            =   7080
         ToolTipText     =   "Zoom descripción"
         Top             =   4035
         Width           =   240
      End
      Begin VB.Label Label29 
         Caption         =   "Observaciones"
         Height          =   255
         Left            =   5880
         TabIndex        =   78
         Top             =   4050
         Width           =   1215
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   1
         Left            =   1080
         ToolTipText     =   "Buscar Situación"
         Top             =   2280
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   0
         Left            =   1080
         ToolTipText     =   "Buscar Colectivo"
         Top             =   1920
         Width           =   240
      End
      Begin VB.Label Label19 
         Caption         =   "Situación"
         Height          =   255
         Left            =   240
         TabIndex        =   67
         Top             =   2280
         Width           =   855
      End
      Begin VB.Label Label18 
         Caption         =   "Colectivo"
         Height          =   255
         Left            =   240
         TabIndex        =   66
         Top             =   1920
         Width           =   735
      End
      Begin VB.Label Label7 
         Caption         =   "Provincia"
         Height          =   255
         Left            =   240
         TabIndex        =   65
         Top             =   1605
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Población"
         Height          =   255
         Index           =   26
         Left            =   240
         TabIndex        =   64
         Top             =   1240
         Width           =   735
      End
      Begin VB.Label Label6 
         Caption         =   "Domicilio"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   63
         Top             =   880
         Width           =   735
      End
      Begin VB.Label Label5 
         Caption         =   "NIF"
         Height          =   255
         Left            =   240
         TabIndex        =   62
         Top             =   525
         Width           =   735
      End
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   330
      Left            =   4275
      Top             =   6930
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
      TabIndex        =   76
      Top             =   0
      Width           =   11715
      _ExtentX        =   20664
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   22
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
            Object.ToolTipText     =   "Nuevo"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Modificar"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Eliminar"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Buscar Tarjeta"
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Imprimir"
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Tarjetas Libres"
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Codigo Cliente Libre"
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Importación de Clientes"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Primero"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Anterior"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button21 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Siguiente"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button22 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Último"
            ImageIndex      =   9
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin VB.CheckBox chkVistaPrevia 
         Caption         =   "Vista previa"
         Height          =   195
         Index           =   0
         Left            =   8520
         TabIndex        =   77
         Top             =   120
         Width           =   1215
      End
   End
   Begin VB.CommandButton cmdRegresar 
      Caption         =   "&Regresar"
      Height          =   375
      Left            =   10440
      TabIndex        =   69
      Top             =   6960
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.Menu mnOpciones 
      Caption         =   "&Opciones"
      Begin VB.Menu mnBuscar 
         Caption         =   "&Buscar"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnVerTodos 
         Caption         =   "&Ver Todos"
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
      Begin VB.Menu mnBuscarTarjeta 
         Caption         =   "Buscar &Tarjeta"
         Shortcut        =   ^T
      End
      Begin VB.Menu mnImprimir 
         Caption         =   "&Imprimir"
         Shortcut        =   ^I
      End
      Begin VB.Menu mnTarjetasLibres 
         Caption         =   "Tarjetas &Libres"
         Shortcut        =   ^L
      End
      Begin VB.Menu mnClientesLibres 
         Caption         =   "Clientes &Libres"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnImportacion 
         Caption         =   "Importación"
         Shortcut        =   ^O
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
Attribute VB_Name = "frmManClien"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' +-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+
' +-+- Autor: MANOLO                   -+-+
' +-+- Menú: CLIENTES                  -+-+
' +-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+

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

'Dim T1 As Single

Public DatosADevolverBusqueda As String    'Tindrà el nº de text que vol que torne, empipat
Public Event DatoSeleccionado(CadenaSeleccion As String)
Public NuevoCodigo As String
Public CodigoActual As String
Public DeConsulta As Boolean

Public NumTarj As Long

' *** declarar els formularis als que vaig a cridar ***
Private WithEvents frmB As frmBuscaGrid
Attribute frmB.VB_VarHelpID = -1
Private WithEvents frmC As frmCal 'calendario fecha
Attribute frmC.VB_VarHelpID = -1
Private WithEvents frmZ As frmZoom  'Zoom para campos Text
Attribute frmZ.VB_VarHelpID = -1
Private WithEvents frmArt As frmManArtic ' articulos
Attribute frmArt.VB_VarHelpID = -1

Private WithEvents frmTra As frmTraerTarje 'Buscar Tarjeta
Attribute frmTra.VB_VarHelpID = -1
Private WithEvents frmFPa As frmManFpago 'F.Pago
Attribute frmFPa.VB_VarHelpID = -1
Private WithEvents frmCoo As frmManCoope 'Colectivos
Attribute frmCoo.VB_VarHelpID = -1
Private WithEvents frmSit As frmManSitua 'Situaciones
Attribute frmSit.VB_VarHelpID = -1
Private WithEvents FrmCli2 As frmManClien2 'Tarjetas
Attribute FrmCli2.VB_VarHelpID = -1

Private WithEvents frmCtas As frmCtasConta 'cuentas contables
Attribute frmCtas.VB_VarHelpID = -1

Private WithEvents frmLis As frmListado 'Listado para impresion de tarjetas
Attribute frmLis.VB_VarHelpID = -1


' *****************************************************


Private Modo As Byte
'*************** MODOS ********************
'   0.-  Formulari net sense cap camp ple
'   1.-  Preparant per a fer la búsqueda
'   2.-  Ja tenim registres i els anem a recorrer
'        ,podem editar-los Edició del camp
'   3.-  Inserció de nou registre
'   4.-  Modificar
'   5.-  Manteniment Llinies

'+-+-Variables comuns a tots els formularis+-+-+

Dim ModoLineas As Byte
'1.- Afegir,  2.- Modificar,  3.- Borrar,  0.-Passar control a Llínies

Dim NumTabMto As Integer 'Indica quin nº de Tab està en modo Mantenimient
Dim TituloLinea As String 'Descripció de la llínia que està en Mantenimient
Dim PrimeraVez As Boolean

Private CadenaConsulta As String 'SQL de la taula principal del formulari
Private Ordenacion As String
Private NombreTabla As String  'Nom de la taula
Private NomTablaLineas As String 'Nom de la Taula de llínies del Mantenimient en que estem

Private kCampo As Integer
'-------------------------------------------------------------------------
Private HaDevueltoDatos As Boolean

'Private VieneDeBuscar As Boolean
'Per a quan torna 2 poblacions en el mateix codi Postal. Si ve de pulsar prismatic
'de búsqueda posar el valor de població seleccionada i no tornar a recuperar de la Base de Datos

Dim btnPrimero As Byte 'Variable que indica el nº del Botó PrimerRegistro en la Toolbar1
'Dim CadAncho() As Boolean  'array, per a quan cridem al form de llínies
Dim indice As Byte 'Index del text1 on es poses els datos retornats des d'atres Formularis de Mtos
Dim CadB As String

Private BuscaChekc As String


'Cambio en cuentas de la contabilidad
Dim IbanAnt As String
Dim NombreAnt As String
Dim BancoAnt  As String
Dim SucurAnt As String
Dim DigitoAnt As String
Dim CuentaAnt As String

Dim DirecAnt As String
Dim cPostalAnt As String
Dim PoblaAnt As String
Dim ProviAnt As String
Dim NifAnt As String

Dim EMaiAnt As String
Dim WebAnt As String




Private Sub btnBuscar_Click(Index As Integer)
     TerminaBloquear
    
    Select Case Index
        Case 0 'Articulo
            
            indice = Index + 15
            
            Set frmArt = New frmManArtic
            frmArt.DatosADevolverBusqueda = "0|1|"
            frmArt.CodigoActual = txtAux(indice).Text
            frmArt.Show vbModal
            Set frmArt = Nothing
            
            PonerFoco txtAux(indice)
    
    End Select
    
    If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Me.Data1, 1
End Sub

Private Sub chkAux_Click(Index As Integer)
    If Modo = 1 Then
        'Buscqueda
        If InStr(1, BuscaChekc, "chkAux(" & Index & ")") = 0 Then BuscaChekc = BuscaChekc & "chkAux(" & Index & ")|"
    End If
End Sub

Private Sub cmbAux_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub cmdAccCRM_Click(Index As Integer)
    
    'Acciones parar el CRM
    Select Case Index
    Case 1
        If Modo <> 2 Then Exit Sub
        If Data1.Recordset.EOF Then Exit Sub
        If text1(0).Text = "" Then Exit Sub
        
        
        frmCRMImprimir.text1 = text1(0).Text
        frmCRMImprimir.text2 = text1(1).Text
        frmCRMImprimir.Show vbModal
        
    Case 0
    
        Select Case CByte(RecuperaValor(lwCRM.Tag, 1))
        Case 0
            'NUEVA, modificar o insertar acciones comerciales
            frmCRMMto.DesdeElCliente = Data1.Recordset!codClien
            frmCRMMto.TipoPredefinido = 0  'sin tipo predefinido
            frmCRMMto.DatosADevolverBusqueda = ""   'NUEVA
            frmCRMMto.Show vbModal
        Case 1
            'NUEVA llamda EFECTUADA
            frmCRMMto.DesdeElCliente = Data1.Recordset!codClien
            frmCRMMto.TipoPredefinido = 1  'Llamada efectuada
            frmCRMMto.DatosADevolverBusqueda = ""   'NUEVA
            frmCRMMto.Show vbModal
            
        Case 2
            'Emails
'            LanzarProgramaEmails
'            If MsgBox("Refrescar datos?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
        Case 3
            'NO puede insertar nada.
            Exit Sub
        Case 4
'            frmCrmObsDpto.Nuevo = True
'            frmCrmObsDpto.Label2.Caption = Data1.Recordset!NomClien
'            frmCrmObsDpto.Tag = Data1.Recordset!CodClien
'            frmCrmObsDpto.Show vbModal
            
        Case 5
            BuscaChekc = ""
            If text1(9).Text = "" Then
                BuscaChekc = "No tiene cta contable"
            Else
                If text2(9).Text = "" Then BuscaChekc = "Cta contable incorrecta"
            End If
            If BuscaChekc < "" Then
                MsgBox BuscaChekc, vbExclamation
                Exit Sub
            End If
            BuscaChekc = "-1|" & text1(1).Text & "|" & text1(9).Text & "|" & text2(9).Text & "|"
            frmCRMReclamas.Intercambio = BuscaChekc  'nueva
            frmCRMReclamas.Show vbModal
            BuscaChekc = ""
        Case 6
            'NUEVA entrada en Historial
            frmCRMMto.DesdeElCliente = Data1.Recordset!codsocio
            frmCRMMto.TipoPredefinido = 2  'Historial
            frmCRMMto.DatosADevolverBusqueda = ""   'NUEVA
            frmCRMMto.Show vbModal
        End Select
        Me.Refresh
        DoEvents
        CargaDatosLWCRM
        Screen.MousePointer = vbDefault
    Case 2
    
        If CByte(RecuperaValor(lwCRM.Tag, 1)) = 4 Then
'            If lwCRM.SelectedItem Is Nothing Then Exit Sub
'            If MsgBox("¿Desea eliminar las observaciones del departamento " & Me.lwCRM.SelectedItem.Text & "?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
'
'            BuscaChekc = "DELETE from scrmobsclien  WHERE codclien = " & Me.Data1.Recordset!CodClien & " AND dpto=" & lwCRM.SelectedItem.SubItems(3)
'            If ejecutar(BuscaChekc, False) Then CargaDatosLWCRM
'            BuscaChekc = ""
        ElseIf CByte(RecuperaValor(lwCRM.Tag, 1)) = 6 Then
        
        End If
    End Select
End Sub


Private Sub cmdAceptar_Click()
    Screen.MousePointer = vbHourglass
    
    On Error GoTo Error1
    
    Select Case Modo
        Case 1  'BÚSQUEDA
            HacerBusqueda
            
        Case 3 'INSERTAR
            If DatosOk Then
                If InsertarDesdeForm2(Me, 1) Then
                    text2(9).Text = PonerNombreCuenta(text1(9), Modo, text1(0).Text)
        
                    Data1.RecordSource = "Select * from " & NombreTabla & Ordenacion
                    PosicionarData
                End If
            Else
                ModoLineas = 0
            End If
            
        Case 4  'MODIFICAR
            If DatosOk Then
                If ModificaDesdeFormulario2(Me, 1) Then
                
                    '[Monica]17/01/2014: Si han cambiado nombre o CCC pregunto si quieren cambiar los datos de la cuenta en la seccion de horto
                    ModificarDatosCuentaContable
                 
                    TerminaBloquear
                    PosicionarData
                End If
            Else
                ModoLineas = 0
            End If
        ' *** si n'hi han llínies ***
        Case 5 'LLÍNIES
            Select Case ModoLineas
                Case 1 'afegir llínia
                    InsertarLinea
                Case 2 'modificar llínies
                    If ModificarLinea Then
                        PosicionarData
                    Else
                        PonerFoco txtAux(12)
                    End If
            End Select
        ' **************************
    End Select
    Screen.MousePointer = vbDefault
 
Error1:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub

Private Sub Combo1_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub Form_Activate()
    Screen.MousePointer = vbDefault
    If PrimeraVez Then PrimeraVez = False
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub Form_Unload(Cancel As Integer)
    CheckValueGuardar Me.Name, Me.chkVistaPrevia(0).Value
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
Dim I As Integer

    PrimeraVez = True
    
    ' ICONETS DE LA BARRA
    btnPrimero = 18 'index del botó "primero"
    With Me.Toolbar1
        .HotImageList = frmPpal.imgListComun_OM
        .DisabledImageList = frmPpal.imgListComun_BN
        .ImageList = frmPpal.imgListComun
        'l'1 i el 2 son separadors
        .Buttons(3).Image = 1   'Buscar
        .Buttons(4).Image = 2   'Totss
        'el 5 i el 6 son separadors
        .Buttons(7).Image = 3   'Insertar
        .Buttons(8).Image = 4   'Modificar
        .Buttons(9).Image = 5   'Borrar
        .Buttons(11).Image = 21   'Borrar
        'el 10 i el 11 son separadors
        .Buttons(12).Image = 10  'Imprimir
        .Buttons(13).Image = 24  ' busqueda de tarjetas libres
        .Buttons(14).Image = 16  ' busqueda del siguiente codigo de cliente libre
        .Buttons(15).Image = 19  'Importacion de clientes
        .Buttons(16).Image = 11  'Eixir
        'el 13 i el 14 son separadors
        .Buttons(btnPrimero).Image = 6  'Primer
        .Buttons(btnPrimero + 1).Image = 7 'Anterior
        .Buttons(btnPrimero + 2).Image = 8 'Següent
        .Buttons(btnPrimero + 3).Image = 9 'Últim
    End With
    
    '[Monica]04/11/2015: para Pobla del Duc, nuevo punto de importacion de socios
    Me.Toolbar1.Buttons(15).visible = (vParamAplic.Cooperativa = 4)
    Me.Toolbar1.Buttons(15).Enabled = (vParamAplic.Cooperativa = 4)
    Me.mnImportacion.visible = (vParamAplic.Cooperativa = 4)
    Me.mnImportacion.Enabled = (vParamAplic.Cooperativa = 4)
    
    
    ' ******* si n'hi han llínies *******
    'ICONETS DE LES BARRES ALS TABS DE LLÍNIA
    For I = 0 To ToolAux.Count - 1
        With Me.ToolAux(I)
            .HotImageList = frmPpal.imgListComun_OM16
            .DisabledImageList = frmPpal.imgListComun_BN16
            .ImageList = frmPpal.imgListComun16
            .Buttons(1).Image = 3   'Insertar
            .Buttons(2).Image = 4   'Modificar
            .Buttons(3).Image = 5   'Borrar
        End With
'        If i = 0 Then
'            Me.ToolAux(i).Buttons(5).Image = 24
'        End If
    Next I
    ' ***********************************
    
    'cargar IMAGES de busqueda
    For I = 0 To Me.imgBuscar.Count - 1
        Me.imgBuscar(I).Picture = frmPpal.imgListImages16.ListImages(1).Picture
    Next I
    
    'carga IMAGES de mail
    For I = 0 To Me.ImgMail.Count - 1
        Me.ImgMail(I).Picture = frmPpal.imgListImages16.ListImages(2).Picture
    Next I
    
    'IMAGES para zoom
    For I = 0 To Me.imgZoom.Count - 1
        Me.imgZoom(I).Picture = frmPpal.imgListImages16.ListImages(3).Picture
    Next I
    
    Me.imgDoc(1).Picture = frmPpal.imgListPpal.ListImages(11).Picture


    Label23.visible = (vParamAplic.Cooperativa = 1)
    Me.imgDoc(1).visible = (vParamAplic.Cooperativa = 1)

    
    ImagenesNavegacion
    
    ' *** si n'hi han tabs, per a que per defecte sempre es pose al 1r***
    Me.SSTab1.Tab = 0
    ' *******************************************************************
    
    LimpiarCampos   'Neteja els camps TextBox
    ' ******* si n'hi han llínies *******
    DataGridAux(0).ClearFields
    DataGridAux(1).ClearFields
    DataGridAux(1).ClearFields
    
    '*** canviar el nom de la taula i l'ordenació de la capçalera ***
    NombreTabla = "ssocio"
    Ordenacion = " ORDER BY codsocio"
    
    'Mirem com està guardat el valor del check
    chkVistaPrevia(0).Value = CheckValueLeer(Name)
    
    Data1.ConnectionString = Conn
    '***** cambiar el nombre de la PK de la cabecera *************
    Data1.RecordSource = "Select * from " & NombreTabla & " where codsocio=-1"
    Data1.Refresh
       
    ModoLineas = 0
       
    ' *** si n'hi han combos (capçalera o llínies) ***
    CargaCombo 0
    
    If DatosADevolverBusqueda = "" Then
        PonerModo 0
    Else
        PonerModo 1 'búsqueda
        ' *** posar de groc els camps visibles de la clau primaria de la capçalera ***
        text1(0).BackColor = vbYellow 'codclien
        ' ****************************************************************************
    End If
    
    Me.SSTab1.TabEnabled(4) = vParamAplic.HayCRM
    
    If vParamAplic.HayCRM Then CargaColumnasCRM 3
    
    
End Sub

Private Sub LimpiarCampos()
Dim I As Integer

    
    On Error Resume Next
    
    limpiar Me   'Mètode general: Neteja els controls TextBox
    lblIndicador.Caption = ""
    
    ' *** si n'hi han combos a la capçalera ***
    Me.Combo1(0).ListIndex = -1
    Me.Combo1(1).ListIndex = -1
    ' *****************************************

    For I = 0 To chkAux.Count - 1
        chkAux(I).Value = 0
    Next I
    
    lwCRM.ListItems.Clear

    '13/02/2007 he tenido que limpiar el combo de lineas pq en algun lado se ha cargado
    Me.cmbAux(0).ListIndex = -1

    If Err.Number <> 0 Then Err.Clear
End Sub

Private Sub LimpiarCamposLin(frameAux As String)
    On Error Resume Next
    
    LimpiarLin Me, frameAux  'Mètode general: Neteja els controls TextBox
    lblIndicador.Caption = ""

    If Err.Number <> 0 Then Err.Clear
End Sub

'----------------------------------------------------------------
'----------------------------------------------------------------
'   En PONERMODO s'habiliten, o no, els diversos camps del
'   formulari en funció del modo en que anem a treballar
Private Sub PonerModo(Kmodo As Byte, Optional indFrame As Integer)
Dim I As Integer, Numreg As Byte
Dim b As Boolean

    On Error GoTo EPonerModo
 
    BuscaChekc = ""
 
    'Actualisa Iconos Insertar,Modificar,Eliminar
    'ActualizarToolbar Modo, Kmodo
    Modo = Kmodo
    PonerIndicador lblIndicador, Modo, ModoLineas
       
    'Modo 2. N'hi han datos i estam visualisant-los
    '=========================================
    'Posem visible, si es formulari de búsqueda, el botó "Regresar" quan n'hi han datos
    If DatosADevolverBusqueda <> "" Then
        cmdRegresar.visible = (Modo = 2)
    Else
        cmdRegresar.visible = False
    End If
    
    '=======================================
    b = (Modo = 2)
    'Posar Fleches de desplasament visibles
    Numreg = 1
    If Not Data1.Recordset.EOF Then
        If Data1.Recordset.RecordCount > 1 Then Numreg = 2 'Només es per a saber que n'hi ha + d'1 registre
    End If
    DesplazamientoVisible Me.Toolbar1, btnPrimero, b, Numreg
    
    '---------------------------------------------
    b = Modo <> 0 And Modo <> 2
    cmdCancelar.visible = b
    cmdAceptar.visible = b
       
    'Bloqueja els camps Text1 si no estem modificant/Insertant Datos
    'Si estem en Insertar a més neteja els camps Text1
    BloquearText1 Me, Modo
    
    '*** si n'hi han combos a la capçalera ***
    BloquearCombo Me, Modo
    '**************************
    
    ' ***** bloquejar tots els controls visibles de la clau primaria de la capçalera ***
'    If Modo = 4 Then _
'        BloquearTxt Text1(0), True 'si estic en  modificar, bloqueja la clau primaria
    ' **********************************************************************************
    
    ' **** si n'hi han imagens de buscar en la capçalera *****
    BloquearImgBuscar Me, Modo, ModoLineas
    BloquearImgZoom Me, Modo, ModoLineas
    BloquearImgFec Me, 15, Modo
    BloquearImgFec Me, 16, Modo
    ' ********************************************************
    
    chkVistaPrevia(0).Enabled = (Modo <= 2)
    
    PonerLongCampos

    If (Modo < 2) Or (Modo = 3) Then
        For I = 0 To 2
            CargaGrid I, False
        Next I
    End If
    
    b = (Modo = 4) Or (Modo = 2)
    For I = 0 To 2
        DataGridAux(I).Enabled = b
    Next I
      
    For I = 0 To chkAux.Count - 1
        chkAux(I).Enabled = (Modo = 1 Or Modo = 3 Or Modo = 4)
    Next I
    
    cmdAccCRM(0).visible = vParamAplic.HayCRM And Modo = 2
    cmdAccCRM(1).visible = vParamAplic.HayCRM And Modo = 2
    
    
    
    ' ****** si n'hi han combos a la capçalera ***********************
    If (Modo = 0) Or (Modo = 2) Or (Modo = 5) Then
        Combo1(0).Enabled = False
        Combo1(0).BackColor = &H80000018 'groc
        Combo1(1).Enabled = False
        Combo1(1).BackColor = &H80000018 'groc
    ElseIf (Modo = 1) Or (Modo = 3) Or (Modo = 4) Then
        Combo1(0).Enabled = True
        Combo1(0).BackColor = &H80000005 'blanc
        Combo1(1).Enabled = True
        Combo1(1).BackColor = &H80000005 'blanc
    End If
    ' ****************************************************************
    
    PonerModoOpcionesMenu (Modo) 'Activar opcions menú según modo
    PonerOpcionesMenu   'Activar opcions de menú según nivell
                        'de permisos de l'usuari


    If Modo <> 2 Then
        lwCRM.ListItems.Clear
        If vParamAplic.HayCRM Then lwCRM.ListItems.Clear
    End If


EPonerModo:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub

Private Sub PonerLongCampos()
'Modificar el MaxLength del campo en funcion de si es modo de búsqueda o no
'para los campos que permitan introducir criterios más largos del tamaño del campo
'Si maxlength=6 para codprove, en modo busqueda Maxlenth=13 'el doble + 1
    
    'para los TEXT1
    PonerLongCamposGnral Me, Modo, 1
End Sub

Private Sub PonerOpcionesMenu()
    PonerOpcionesMenuGeneral Me
    PonerOpcionesMenuGeneralNew Me
End Sub

Private Sub PonerModoOpcionesMenu(Modo)
'Actives unes Opcions de Menú i Toolbar según el modo en que estem
Dim b As Boolean, bAux As Boolean
Dim I As Byte
    
    'Barra de CAPÇALERA
    '------------------------------------------
    'b = (Modo = 2 Or Modo = 0 Or Modo = 1)
    b = (Modo = 2 Or Modo = 0)
    'Buscar
    Toolbar1.Buttons(3).Enabled = b
    Me.mnBuscar.Enabled = b
    'Vore Tots
    Toolbar1.Buttons(4).Enabled = b
    Me.mnVerTodos.Enabled = b
    
    'Insertar
    Toolbar1.Buttons(7).Enabled = b And Not DeConsulta
    Me.mnNuevo.Enabled = b And Not DeConsulta
    
    b = (Modo = 2 And Data1.Recordset.RecordCount > 0) And Not DeConsulta
    'Modificar
    Toolbar1.Buttons(8).Enabled = b
    Me.mnModificar.Enabled = b
    'eliminar
    Toolbar1.Buttons(9).Enabled = b
    Me.mnEliminar.Enabled = b
    
    'Imprimir
    'Toolbar1.Buttons(12).Enabled = (b Or Modo = 0)
    Toolbar1.Buttons(12).Enabled = True And Not DeConsulta
       
    ' *** si n'hi han llínies que tenen grids (en o sense tab) ***
    b = (Modo = 3 Or Modo = 4 Or Modo = 2) And Not DeConsulta
    For I = 0 To ToolAux.Count - 1
        ToolAux(I).Buttons(1).Enabled = b
        If b Then bAux = (b And Me.AdoAux(I).Recordset.RecordCount > 0)
        ToolAux(I).Buttons(2).Enabled = bAux
        ToolAux(I).Buttons(3).Enabled = bAux
    Next I
    
End Sub

Private Sub Desplazamiento(Index As Integer)
'Botons de Desplaçament; per a desplaçar-se pels registres de control Data
    If Data1.Recordset.EOF Then Exit Sub
    DesplazamientoData Data1, Index
    PonerCampos
End Sub

Private Function MontaSQLCarga(Index As Integer, enlaza As Boolean) As String
'--------------------------------------------------------------------
' MontaSQlCarga:
'   Basant-se en la informació proporcionada pel vector de camps
'   crea un SQl per a executar una consulta sobre la base de datos que els
'   torne.
' Si ENLAZA -> Enlaça en el data1
'           -> Si no el carreguem sense enllaçar a cap camp
'--------------------------------------------------------------------
Dim SQL As String
Dim tabla As String
    
    ' ********* si n'hi han tabs, dona igual si en datagrid o no ***********
    Select Case Index
               
        Case 0 'TARJETAS
            SQL = "SELECT codsocio,numlinea,numtarje,tiptarje,CASE tiptarje WHEN 0 THEN ""Normal"" WHEN 1 THEN ""Bonificado"" WHEN 2 THEN ""Profesional"" END, nomtarje,iban, codbanco,codsucur,digcontr,cuentaba, matricul "
            SQL = SQL & " FROM starje "
            If enlaza Then
                SQL = SQL & ObtenerWhereCab(True)
            Else
                SQL = SQL & " WHERE starje.codsocio = -1"
            End If
            SQL = SQL & " ORDER BY starje.numlinea"
               
        Case 1 'MATRICULAS
            SQL = "SELECT codsocio,numlinea,matricul,observac "
            SQL = SQL & " FROM smatri"
            If enlaza Then
                SQL = SQL & ObtenerWhereCab(True)
            Else
                SQL = SQL & " WHERE smatri.codsocio = -1"
            End If
            SQL = SQL & " ORDER BY smatri.numlinea"
            
        Case 2 'MARGENES DE PRECIOS DE ARTICULOS DE COMBUSTIBLE
            SQL = "SELECT codsocio,numlinea,smargen.codartic, nomartic, margen, euroslitro "
            SQL = SQL & " FROM smargen INNER JOIN sartic ON smargen.codartic = sartic.codartic "
            If enlaza Then
                SQL = SQL & ObtenerWhereCab(True)
            Else
                SQL = SQL & " WHERE smargen.codsocio = -1"
            End If
            SQL = SQL & " ORDER BY smargen.numlinea"
            
            
            
    End Select
    
    MontaSQLCarga = SQL
End Function

Private Sub frmArt_DatoSeleccionado(CadenaSeleccion As String)
    txtAux(indice).Text = RecuperaValor(CadenaSeleccion, 1)
    FormateaCampo txtAux(indice)
    text2(7).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmB_Selecionado(CadenaDevuelta As String)
    Dim Aux As String
    
    If CadenaDevuelta <> "" Then
        HaDevueltoDatos = True
        Screen.MousePointer = vbHourglass
        'Sabem quins camps son els que mos torna
        'Creem una cadena consulta i posem els datos
        CadB = ""
        Aux = ValorDevueltoFormGrid(text1(0), CadenaDevuelta, 1)
        CadB = Aux
        '   Com la clau principal es única, en posar el sql apuntant
        '   al valor retornat sobre la clau ppal es suficient
        CadenaConsulta = "select * from " & NombreTabla & " WHERE " & CadB & " " & Ordenacion
        ' **********************************
        PonerCadenaBusqueda
        Screen.MousePointer = vbDefault
    End If
End Sub

Private Sub frmTra_Actualizar(vValor As Integer)
'Mantenimiento de Colectivos
    
    LimpiarCampos
    text1(0).Text = vValor 'codcoope
    
    FormateaCampo text1(0)
'    text2(7).Text = RecuperaValor(CadenaSeleccion, 2) 'nomcoope
        Modo = 1
        cmdAceptar_Click
End Sub

Private Sub frmZ_Actualizar(vCampo As String)
     text1(indice).Text = vCampo
End Sub

Private Sub imgDoc_Click(Index As Integer)
Dim SQL As String

    If Me.AdoAux(0).Recordset.EOF Then Exit Sub
    TerminaBloquear
    
    Select Case Index
        Case 1 'documentos de alta baja de socios/campos
            Set frmLis = New frmListado
            frmLis.Socio = text1(0).Text
            frmLis.Tarjeta = Me.AdoAux(0).Recordset!Numtarje
            frmLis.OpcionListado = 16
            frmLis.Show vbModal
            Set frmLis = Nothing
    End Select

End Sub

' *** si n'hi ha buscar data, posar a les <=== el menor index de les imagens de buscar data ***
' NOTA: ha de coincidir l'index de la image en el del camp a on va a parar el valor
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
    If text1(Index).Text <> "" Then frmC.NovaData = text1(Index).Text
    ' ********************************************

    frmC.Show vbModal
    Set frmC = Nothing
    ' *** repasar si el camp es txtAux o Text1 ***
    PonerFoco text1(CByte(imgFec(15).Tag)) '<===
    ' ********************************************
End Sub

Private Sub frmC_Selec(vFecha As Date)
    ' *** repasar si el camp es txtAux o Text1 ***
    text1(CByte(imgFec(15).Tag)).Text = Format(vFecha, "dd/mm/yyyy") '<===
    ' ********************************************
End Sub
' *****************************************************

'Private Sub btnFec_Click(Index As Integer)
'    imgFec_Click (Index)
'End Sub

Private Sub imgMail_Click(Index As Integer)
    If Index = 0 Then
        If text1(14).Text <> "" Then
            LanzaMailGnral text1(14).Text
        End If
    End If
End Sub

Private Sub imgZoom_Click(Index As Integer)
    
    Set frmZ = New frmZoom

    If Index = 0 Then
        indice = 24
        frmZ.pTitulo = "Observaciones del Cliente"
        frmZ.pValor = text1(indice).Text
        frmZ.pModo = Modo
    
        frmZ.Show vbModal
        Set frmZ = Nothing
            
        PonerFoco text1(indice)
    End If
End Sub

Private Sub lwCRM_DblClick()
Dim clave As String
Dim I As Integer
    
    If Modo <> 2 Then Exit Sub
    If lwCRM.ListItems.Count = 0 Then Exit Sub
    If lwCRM.SelectedItem Is Nothing Then Exit Sub

    'Llegados aqui
    Select Case CByte(RecuperaValor(lwCRM.Tag, 1))
    Case 0
        'Aciones comerciales
        ' modificar o insertar acciones comerciales
        frmCRMMto.DesdeElCliente = Data1.Recordset!codClien
        frmCRMMto.TipoPredefinido = 0  'sin tipo predefinido
        frmCRMMto.DatosADevolverBusqueda = "`fechora`=" & DBSet(lwCRM.SelectedItem.Text, "FH") & _
            " AND scrmacciones.Tipo = " & lwCRM.SelectedItem.SubItems(4) & " And codClien = " & Data1.Recordset!codClien
        frmCRMMto.Show vbModal
    Case 1
        'Llamadas
        If lwCRM.SelectedItem.SmallIcon = 27 Then
'            'Lee de sllama
'
'            CadenaDesdeOtroForm = "`feholla`=" & DBSet(lwCRM.SelectedItem.Text, "FH") & " and `usuario`=" & DBSet(lwCRM.SelectedItem.SubItems(1), "T")
'            frmLLamadasDatos2.SoloVer = True
'            frmLLamadasDatos2.vModo = 4
'            frmLLamadasDatos2.Show vbModal
        Else
            'Lee de acciones realizadas con tipo=1 .....
            
            frmCRMMto.DesdeElCliente = Data1.Recordset!codClien
            frmCRMMto.TipoPredefinido = 1 'Llamadas realizadas
            frmCRMMto.DatosADevolverBusqueda = "`fechora`=" & DBSet(lwCRM.SelectedItem.Text, "FH") & " AND scrmacciones.Tipo = 1 And codClien = " & Data1.Recordset!codClien
            frmCRMMto.Show vbModal
            
        End If
    Case 2
        'MAIL
        frmMensajes.OpcionMensaje = 21
        If lwCRM.SelectedItem.SmallIcon = 28 Then
            frmMensajes.cadWHERE2 = "0"
        Else
            frmMensajes.cadWHERE2 = "1"
        End If
        frmMensajes.cadWhere = "codclien = " & text1(0).Text & " AND  entryID = '" & lwCRM.SelectedItem.SubItems(5) & "'"
        frmMensajes.Show vbModal
    Case 3
        'Cobros. NO HACEMOS NADA
        'Nos piramos
        Exit Sub
        
    Case 4
'        frmCrmObsDpto.Nuevo = False
'        BuscaChekc = "dpto = " & Me.lwCRM.SelectedItem.SubItems(3) & " AND codclien "
'        CadenaDesdeOtroForm = DevuelveDesdeBD(conAri, "observa", "scrmobsclien", BuscaChekc, CStr(Data1.Recordset!codClien))
'
'        frmCrmObsDpto.Dpto = CByte(Me.lwCRM.SelectedItem.SubItems(3))
'        frmCrmObsDpto.Label2.Caption = Data1.Recordset!NomClien
'        frmCrmObsDpto.Tag = Data1.Recordset!codClien
'        frmCrmObsDpto.Show vbModal
'
    Case 5
        'Reclamas n
            BuscaChekc = lwCRM.SelectedItem.SubItems(4) & "|" & text1(1).Text & "|"
            frmCRMReclamas.Intercambio = BuscaChekc
            frmCRMReclamas.Show vbModal
    
    Case 6
            frmCRMMto.DesdeElCliente = Data1.Recordset!codsocio
            frmCRMMto.TipoPredefinido = 2 'Historial
            frmCRMMto.DatosADevolverBusqueda = "`fechora`=" & DBSet(lwCRM.SelectedItem.Text, "FH") & " AND scrmacciones.Tipo = 2 And codClien = " & Data1.Recordset!codsocio
            frmCRMMto.Show vbModal
    End Select
    Me.Refresh
    DoEvents
    
    
    If CByte(RecuperaValor(lwCRM.Tag, 1)) = 5 Then
        clave = lwCRM.SelectedItem.SubItems(4)
    Else
        clave = lwCRM.SelectedItem.Text
    End If
    CargaDatosLWCRM
    
    Set lwCRM.SelectedItem = Nothing
    If CByte(RecuperaValor(lwCRM.Tag, 1)) = 5 Then
        'para encontrar en las reclamas debe buscar por el campo codigo 4
        For I = 1 To lwCRM.ListItems.Count
            If clave = lwCRM.ListItems(I).SubItems(4) Then
                Set lwCRM.SelectedItem = lwCRM.ListItems(I)
            Else
                lwCRM.ListItems(I).Selected = False
            End If
        Next
    Else
        For I = 1 To lwCRM.ListItems.Count
            If clave = lwCRM.ListItems(I).Text Then
                Set lwCRM.SelectedItem = lwCRM.ListItems(I)
            Else
                lwCRM.ListItems(I).Selected = False
            End If
        Next
    End If
    BuscaChekc = ""

End Sub

Private Sub mnBuscar_Click()
Dim I As Integer

    BotonBuscar
    Combo1(0).ListIndex = -1 'quan busque, per defecte no seleccione cap tipo de client
    Combo1(1).ListIndex = -1
    
    For I = 0 To chkAux.Count - 1
        chkAux(I).Value = 0
    Next I
    
End Sub

Private Sub mnBuscarTarjeta_Click()
    Set frmTra = New frmTraerTarje
    frmTra.DatosADevolverBusqueda = "0|1|"
    frmTra.CodigoActual = text1(0).Text
    frmTra.Show vbModal
    Set frmTra = Nothing
    PonerFoco text1(0)
End Sub

Private Sub mnClientesLibres_Click()
    BotonClientesLibres
End Sub

Private Sub mnEliminar_Click()
    BotonEliminar
End Sub

Private Sub mnImportacion_Click()
    If Not (Modo = 0 Or Modo = 2) Then Exit Sub
    
    frmTrasSocios.Show vbModal
End Sub

Private Sub mnImprimir_Click()
    AbrirListado (10)
End Sub

Private Sub mnModificar_Click()
    ' *** repasar el nom de l'adodc, l'index del Field i el camp que te la PK ***
    ' ### [Monica] 26/09/2006 dejamos modificar y eliminar el codigo 0
    'El registre de codi 0 no es pot Modificar ni Eliminar
    'If EsCodigoCero(CStr(Data1.Recordset.Fields(0).Value), FormatoCampo(Text1(0))) Then Exit Sub
    ' ***************************************************************************
    
    If BLOQUEADesdeFormulario2(Me, Data1, 1) Then BotonModificar
End Sub

Private Sub mnNuevo_Click()
    BotonAnyadir
End Sub

Private Sub mnSalir_Click()
    Unload Me
End Sub

Private Sub mnTarjetasLibres_Click()
    BotonTarjetasLibres
End Sub

Private Sub mnVerTodos_Click()
    BotonVerTodos
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

    Select Case Button.Index
        Case 3  'Búscar
           mnBuscar_Click
        Case 4  'Tots
            mnVerTodos_Click
        Case 7  'Nou
            mnNuevo_Click
        Case 8  'Modificar
            mnModificar_Click
        Case 9  'Borrar
            mnEliminar_Click
        Case 11  'Buscar Tarjeta
            mnBuscarTarjeta_Click
        Case 12 'Imprimir
            mnImprimir_Click
        Case 13 'Buscar nro de tarjetas libres
            mnTarjetasLibres_Click
        Case 14 ' Buscar siguiente codigo cliente libre
            mnClientesLibres_Click
        Case 15 ' Importacion de clientes
            mnImportacion_Click
        Case 16   'Eixir
            mnSalir_Click
            
        Case btnPrimero To btnPrimero + 3 'Fleches Desplaçament
            Desplazamiento (Button.Index - btnPrimero)
    End Select
End Sub

Private Sub BotonBuscar()
Dim I As Integer
' ***** Si la clau primaria de la capçalera no es Text1(0), canviar-ho en <=== *****
    If Modo <> 1 Then
        LimpiarCampos
        PonerModo 1
        PonerFoco text1(0) ' <===
        text1(0).BackColor = vbYellow ' <===
        ' *** si n'hi han combos a la capçalera ***
        For I = 0 To Combo1.Count - 1
            Combo1(I).ListIndex = -1
        Next I
    Else
        HacerBusqueda
        If Data1.Recordset.EOF Then
            text1(kCampo).Text = ""
            text1(kCampo).BackColor = vbYellow
            PonerFoco text1(kCampo)
        End If
    End If
' ******************************************************************************
End Sub

Private Sub HacerBusqueda()

    CadB = ObtenerBusqueda2(Me, BuscaChekc, 1)
    
    If chkVistaPrevia(0) = 1 Then
        MandaBusquedaPrevia CadB
    ElseIf CadB <> "" Then
        CadenaConsulta = "select * from " & NombreTabla & " WHERE " & CadB & " " & Ordenacion
        PonerCadenaBusqueda
    Else
        ' *** foco al 1r camp visible de la capçalera que siga clau primaria ***
        PonerFoco text1(0)
        ' **********************************************************************
    End If
End Sub

Private Sub MandaBusquedaPrevia(CadB As String)
    Dim cad As String
        
    'Cridem al form
    ' **************** arreglar-ho per a vore lo que es desije ****************
    ' NOTA: el total d'amples de ParaGrid, ha de sumar 100
    cad = ""
    cad = cad & ParaGrid(text1(0), 15, "Cód.")
    cad = cad & ParaGrid(text1(1), 60, "Nombre")
    cad = cad & ParaGrid(text1(2), 25, "N.I.F.")
    If cad <> "" Then
        Screen.MousePointer = vbHourglass
        Set frmB = New frmBuscaGrid
        frmB.vCampos = cad
        frmB.vTabla = NombreTabla
        frmB.vSql = CadB
        HaDevueltoDatos = False
        frmB.vDevuelve = "0|1|2|" '*** els camps que volen que torne ***
        frmB.vTitulo = "Clientes" ' ***** repasa açò: títol de BuscaGrid *****
        frmB.vSelElem = 1

        frmB.Show vbModal
        Set frmB = Nothing
        'Si ha posat valors i tenim que es formulari de búsqueda llavors
        'tindrem que tancar el form llançant l'event
        If HaDevueltoDatos Then
            If (Not Data1.Recordset.EOF) And DatosADevolverBusqueda <> "" Then _
                cmdRegresar_Click
        Else   'de ha retornat datos, es a decir NO ha retornat datos
            PonerFoco text1(kCampo)
        End If
    End If
End Sub

Private Sub cmdRegresar_Click()
Dim cad As String
Dim Aux As String
Dim I As Integer
Dim J As Integer

    If Data1.Recordset.EOF Then
        MsgBox "Ningún registro devuelto.", vbExclamation
        Exit Sub
    End If
    
    cad = ""
    I = 0
    Do
        J = I + 1
        I = InStr(J, DatosADevolverBusqueda, "|")
        If I > 0 Then
            Aux = Mid(DatosADevolverBusqueda, J, I - J)
            J = Val(Aux)
            cad = cad & text1(J).Text & "|"
        End If
    Loop Until I = 0
    RaiseEvent DatoSeleccionado(cad)
    Unload Me
End Sub

Private Sub PonerCadenaBusqueda()
    Screen.MousePointer = vbHourglass
    On Error GoTo EEPonerBusq
    
    Data1.RecordSource = CadenaConsulta
    Data1.Refresh
    
    If Data1.Recordset.RecordCount <= 0 Then
        MsgBox "No hay ningún registro en la tabla " & NombreTabla, vbInformation
        Screen.MousePointer = vbDefault
        Exit Sub
    Else
        'Data1.Recordset.MoveLast
        Data1.Recordset.MoveFirst
        
        PonerModo 2
        
        PonerCampos
    End If
    
    Screen.MousePointer = vbDefault
    Exit Sub

EEPonerBusq:
    MuestraError Err.Number, "PonerCadenaBusqueda"
    PonerModo 0
    Screen.MousePointer = vbDefault
End Sub

Private Sub BotonVerTodos()
'Vore tots
    LimpiarCampos 'Neteja els Text1
    CadB = ""
    
    If chkVistaPrevia(0).Value = 1 Then
        MandaBusquedaPrevia ""
    Else
        CadenaConsulta = "Select * from " & NombreTabla & Ordenacion
        PonerCadenaBusqueda
    End If
    
End Sub

Private Sub BotonAnyadir()

    LimpiarCampos 'Huida els TextBox
    PonerModo 3
    
    ' ****** Valors per defecte a l'afegir, repasar si n'hi ha
    ' codEmpre i quins camps tenen la PK de la capçalera *******
    text1(0).Text = SugerirCodigoSiguienteStr("ssocio", "codsocio")
    FormateaCampo text1(0)
       
    text1(15).Text = Format(Now, "dd/mm/yyyy") ' Quan afegixc pose en F.Alta i F.Modificación la data actual
    PosicionarCombo Combo1(0), 0
    PosicionarCombo Combo1(1), 0
    PosicionarCombo Combo1(2), 0
    PosicionarCombo Combo1(3), 0
    PosicionarCombo Combo1(4), 0
        
    PonerFoco text1(0) '*** 1r camp visible que siga PK ***
    
    ' *** si n'hi han camps de descripció a la capçalera ***
    'PosarDescripcions

    ' *** si n'hi han tabs, em posicione al 1r ***
    Me.SSTab1.Tab = 0
End Sub

Private Sub BotonModificar()

    PonerModo 4
    
    '[Monica]17/01/2014:me guardo los valores de nombre y CCC por si cambian
    NombreAnt = text1(1).Text
    IbanAnt = text1(32).Text
    BancoAnt = text1(18).Text
    SucurAnt = text1(19).Text
    DigitoAnt = text1(20).Text
    CuentaAnt = text1(21).Text
    
    DirecAnt = text1(3).Text
    cPostalAnt = text1(4).Text
    PoblaAnt = text1(5).Text
    ProviAnt = text1(6).Text
    NifAnt = text1(2).Text
    EMaiAnt = text1(14).Text
    WebAnt = text1(10).Text

    ' *** bloquejar els camps visibles de la clau primaria de la capçalera ***
    BloquearTxt text1(0), True
    
    ' *** foco al 1r camp visible que NO siga clau primaria ***
    PonerFoco text1(1)
End Sub

Private Sub BotonEliminar()
Dim cad As String

    On Error GoTo EEliminar

    'Ciertas comprobaciones
    If Data1.Recordset.EOF Then Exit Sub

    ' *** repasar el nom de l'adodc, l'index del Field i el camp que te la PK ***
    ' ### [Monica] 26/09/2006 dejamos modificar y eliminar el codigo 0
    'El registre de codi 0 no es pot Modificar ni Eliminar
    'If EsCodigoCero(CStr(Data1.Recordset.Fields(0).Value), FormatoCampo(text1(0))) Then Exit Sub
    ' ***************************************************************************

    ' *************** canviar la pregunta ****************
    cad = "¿Seguro que desea eliminar el Cliente?"
    cad = cad & vbCrLf & "Código: " & Format(Data1.Recordset.Fields(0), FormatoCampo(text1(0)))
    cad = cad & vbCrLf & "Nombre: " & Data1.Recordset.Fields(2)
    
    If MsgBox(cad, vbQuestion + vbYesNo) = vbYes Then
        On Error GoTo EEliminar
        Screen.MousePointer = vbHourglass
        NumRegElim = Data1.Recordset.AbsolutePosition
        If Not eliminar Then
            Screen.MousePointer = vbDefault
            Exit Sub
        ElseIf SituarDataTrasEliminar(Data1, NumRegElim) Then
            PonerCampos
        Else
            LimpiarCampos
            PonerModo 0
        End If
    End If
    Screen.MousePointer = vbDefault
    
EEliminar:
    Screen.MousePointer = vbDefault
    If Err.Number > 0 Then MuestraError Err.Number, "Eliminar Proveedor", Err.Description
End Sub

Private Sub PonerCampos()
Dim I As Integer
Dim codPobla As String, desPobla As String
Dim CPostal As String, desProvi As String, desPais As String

    If Data1.Recordset.EOF Then Exit Sub
    PonerCamposForma2 Me, Data1, 1 'opcio=1: posa el format o els camps de la capçalera
    
    ' *** si n'hi han llínies en datagrids ***
    'For i = 0 To DataGridAux.Count - 1
    For I = 0 To 2
        CargaGrid I, True
        If Not AdoAux(I).Recordset.EOF Then _
            PonerCamposForma2 Me, AdoAux(I), 2, "FrameAux" & I
    Next I

    ' ************* configurar els camps de les descripcions de la capçalera *************
    text2(7).Text = PonerNombreDeCod(text1(7), "scoope", "nomcoope", "codcoope", "N")
    text2(8).Text = PonerNombreDeCod(text1(8), "ssitua", "nomsitua")
    text2(17).Text = PonerNombreDeCod(text1(17), "sforpa", "nomforpa", "codforpa", "N")
    text2(9).Text = PonerNombreCuenta(text1(9), Modo)
    ' ********************************************************************************
    
    '-- Esto permanece para saber donde estamos
    lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
    
    
    Me.SSTab1.TabEnabled(3) = (CInt(Data1.Recordset!bonifesp) = 1)
    
    
    PonerDatosSegurosCuenta
    
    
'    lblIndicador.Caption = "Datos navegacion"
    Me.Refresh
    DoEvents
    If vParamAplic.HayCRM Then CargaDatosLWCRM
    
    PonerModoOpcionesMenu (Modo)
    PonerOpcionesMenu
End Sub

Private Sub PonerDatosSegurosCuenta()
Dim SQL As String
Dim RS As ADODB.Recordset

    On Error GoTo ePonerDatosSegurosCuenta
    
    If text1(9).Text = "" Then Exit Sub

    SQL = "select numpoliz, fecbajcre, credisol, fecconce, credicon from cuentas where codmacta = " & DBSet(text1(9).Text, "T")
    Set RS = New ADODB.Recordset
    RS.Open SQL, ConnConta, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    text1(26).Text = ""
    text1(27).Text = ""
    text1(29).Text = ""
    text1(30).Text = ""
    
    If Not RS.EOF Then
        text1(26).Text = DBLet(RS.Fields(0).Value)
        text1(27).Text = Format(DBLet(RS.Fields(1).Value), "dd/mm/yyyy")
        text1(29).Text = Format(DBLet(RS.Fields(3).Value), "dd/mm/yyyy")
        text1(30).Text = Format(DBLet(RS.Fields(4).Value), "##,###,###,##0.00")
    End If
    
    Set RS = Nothing
    Exit Sub
    
ePonerDatosSegurosCuenta:
    MuestraError Err.Number, "Poner Datos Seguros", Err.Description
End Sub


Private Sub cmdCancelar_Click()
Dim I As Integer
Dim V

    Select Case Modo
        Case 1, 3 'Búsqueda, Insertar
                LimpiarCampos
                If Data1.Recordset.EOF Then
                    PonerModo 0
                Else
                    PonerModo 2
                    PonerCampos
                End If
                ' *** foco al primer camp visible de la capçalera ***
                PonerFoco text1(0)

        Case 4  'Modificar
                TerminaBloquear
                PonerModo 2
                PonerCampos
                ' *** primer camp visible de la capçalera ***
                PonerFoco text1(0)
        
        Case 5 'LLÍNIES
            Select Case ModoLineas
                Case 1 'afegir llínia
                    ModoLineas = 0
                    ' *** les llínies que tenen datagrid (en o sense tab) ***
                    If NumTabMto = 0 Or NumTabMto = 1 Or NumTabMto = 2 Or NumTabMto = 4 Then
                        DataGridAux(NumTabMto).AllowAddNew = False
                        ' **** repasar si es diu Data1 l'adodc de la capçalera ***
                        'If BLOQUEADesdeFormulario2(Me, Data1, 1) Then BotonModificar 'Modificar
                        LLamaLineas NumTabMto, ModoLineas 'ocultar txtAux
                        DataGridAux(NumTabMto).Enabled = True
                        DataGridAux(NumTabMto).SetFocus

                        ' *** si n'hi han camps de descripció dins del grid, els neteje ***
                        'txtAux2(2).text = ""

                    End If
                    
                    ' *** si n'hi han tabs ***
                    SituarTab (NumTabMto + 1)
                    
                    If Not AdoAux(NumTabMto).Recordset.EOF Then
                        AdoAux(NumTabMto).Recordset.MoveFirst
                    End If

                Case 2 'modificar llínies
                    ModoLineas = 0
                    
                    ' *** si n'hi han tabs ***
                    SituarTab (NumTabMto + 1)
                    LLamaLineas NumTabMto, ModoLineas 'ocultar txtAux
                    PonerModo 4
                    If Not AdoAux(NumTabMto).Recordset.EOF Then
                        ' *** l'Index de Fields es el que canvie de la PK de llínies ***
                        V = AdoAux(NumTabMto).Recordset.Fields(1) 'el 2 es el nº de llinia
                        AdoAux(NumTabMto).Recordset.Find (AdoAux(NumTabMto).Recordset.Fields(1).Name & " =" & V)
                        ' ***************************************************************
                    End If

            End Select
            
            PosicionarData
            
            ' *** si n'hi han llínies en grids i camps fora d'estos ***
            If Not AdoAux(NumTabMto).Recordset.EOF Then
                DataGridAux_RowColChange NumTabMto, 1, 1
            Else
                LimpiarCamposFrame NumTabMto
            End If
    End Select
    
'[Monica]04/11/2015: si dabamos a cancelar se salia del programa de clientes
'    Unload Me
    
End Sub

Private Function DatosOk() As Boolean
Dim b As Boolean
Dim cta As String
Dim cadMen As String

'Dim Datos As String

    On Error GoTo EDatosOK

    DatosOk = False
    b = CompForm2(Me, 1)
    If Not b Then Exit Function
    
    ' *** canviar els arguments de la funcio, el mensage i repasar si n'hi ha codEmpre ***
    If (Modo = 3) Then 'insertar
        'comprobar si existe ya el cod. del campo clave primaria
        If ExisteCP(text1(0)) Then b = False
    End If
    '[Monica]26/03/2012: obligatorio el email si está marcado
    If b And (Modo = 3 Or Modo = 4) Then
        If chkAux(0).Value = 1 Then
            If Trim(text1(14).Text) = "" Then
                MsgBox "Si el cliente tiene marcado Envio de Factura por EMail, debe introducir el email. Revise.", vbExclamation
                PonerFoco text1(14)
                b = False
            End If
        End If
    End If
    
    
    If b And (Modo = 3 Or Modo = 4) Then
        '[Monica]22/11/2013: añadida la comprobacion de que la cuenta contable sea correcta
        If text1(18).Text = "" Or text1(19).Text = "" Or text1(20).Text = "" Or text1(21).Text = "" Then
            '[Monica]20/11/2013: añadido el codigo de iban
            text1(32).Text = ""
            text1(18).Text = ""
            text1(19).Text = ""
            text1(20).Text = ""
            text1(21).Text = ""
        Else
            cta = Format(text1(18).Text, "0000") & Format(text1(19).Text, "0000") & Format(text1(20).Text, "00") & Format(text1(21).Text, "0000000000")
            If Val(ComprobarCero(cta)) = 0 Then
                cadMen = "El cliente no tiene asignada cuenta bancaria."
                MsgBox cadMen, vbExclamation
            End If
            If Not Comprueba_CC(cta) Then
                cadMen = "La cuenta bancaria del cliente no es correcta. ¿ Desea continuar ?."
                If MsgBox(cadMen, vbQuestion + vbYesNo + vbDefaultButton1) = vbYes Then
                    b = True
                Else
                    PonerFoco text1(18)
                    b = False
                End If
            Else
'                '[Monica]20/11/2013: añadimos el tema de la comprobacion del IBAN
'                If Not Comprueba_CC_IBAN(cta, Text1(42).Text) Then
'                    cadMen = "La cuenta IBAN del cliente no es correcta. ¿ Desea continuar ?."
'                    If MsgBox(cadMen, vbQuestion + vbYesNo + vbDefaultButton1) = vbYes Then
'                        b = True
'                    Else
'                        PonerFoco Text1(42)
'                        b = False
'                    End If
'                End If

'       sustituido por lo de David
                BuscaChekc = ""
                If Me.text1(32).Text <> "" Then BuscaChekc = Mid(text1(32).Text, 1, 2)
                    
                If DevuelveIBAN2(BuscaChekc, cta, cta) Then
                    If Me.text1(32).Text = "" Then
                        If MsgBox("Poner IBAN ?", vbQuestion + vbYesNo) = vbYes Then Me.text1(32).Text = BuscaChekc & cta
                    Else
                        If Mid(text1(32).Text, 3) <> cta Then
                            cta = "Calculado : " & BuscaChekc & cta
                            cta = "Introducido: " & Me.text1(32).Text & vbCrLf & cta & vbCrLf
                            cta = "Error en codigo IBAN" & vbCrLf & cta & "Continuar?"
                            If MsgBox(cta, vbQuestion + vbYesNo) = vbNo Then
                                PonerFoco text1(32)
                                b = False
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End If
    
    
    
    ' ************************************************************************************
    DatosOk = b
    
EDatosOK:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Function

Private Sub PosicionarData()
Dim cad As String, Indicador As String

    ' *** canviar-ho per tota la PK de la capçalera, no llevar els () ***
    cad = "(codsocio=" & text1(0).Text & ")"
    
    ' *** gastar SituarData o SituarDataMULTI depenent de si la PK es simple o composta ***
    'If SituarDataMULTI(Data1, cad, Indicador) Then
    If SituarData(Data1, cad, Indicador) Then
        If ModoLineas <> 1 Then
            PonerModo 2
            PonerCampos
        End If
        lblIndicador.Caption = Indicador
    Else
       LimpiarCampos
       PonerModo 0
    End If
End Sub

Private Function eliminar() As Boolean
Dim vWhere As String

    On Error GoTo FinEliminar

    Conn.BeginTrans
    ' ***** canviar el nom de la PK de la capçalera, repasar codEmpre *******
    vWhere = " WHERE codsocio=" & Data1.Recordset!codsocio
        
    ' ***** elimina les llínies ****
    Conn.Execute "DELETE FROM starje " & vWhere
        
    Conn.Execute "DELETE FROM smatri " & vWhere
        
    Conn.Execute "DELETE FROM smargen " & vWhere
        
    'Eliminar la CAPÇALERA
    vWhere = " WHERE codsocio=" & Data1.Recordset!codsocio
    Conn.Execute "Delete from " & NombreTabla & vWhere
       
FinEliminar:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Eliminar"
        Conn.RollbackTrans
        eliminar = False
    Else
        Conn.CommitTrans
        eliminar = True
    End If
End Function

Private Sub Text1_GotFocus(Index As Integer)
    kCampo = Index
    ConseguirFoco text1(Index), Modo
End Sub

' *** si n'hi han combos a la capçalera ***
Private Sub Combo1_GotFocus(Index As Integer)
    If Modo = 1 Then Combo1(Index).BackColor = vbYellow
End Sub

Private Sub Combo1_LostFocus(Index As Integer)
    If Combo1(Index).BackColor = vbYellow Then Combo1(Index).BackColor = vbWhite
End Sub

Private Sub Text1_LostFocus(Index As Integer)
Dim cadMen As String
Dim Nuevo As Boolean

    If Not PerderFocoGnral(text1(Index), Modo) Then Exit Sub
    
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub
    
    ' ***************** configurar els LostFocus dels camps de la capçalera *****************
    Select Case Index
        Case 0 'cod cliente
            PonerFormatoEntero text1(0)

        Case 1 'NOMBRE
            text1(Index).Text = UCase(text1(Index).Text)
        
        Case 2 'NIF
            text1(Index).Text = UCase(text1(Index).Text)
            ValidarNIF text1(Index).Text
                
        Case 10, 11, 12, 13, 14 'telèfons, fax i mòbils
'            PosarFormatTelefon Text1(Index)
                
        Case 7 'COLECTIVO
            If PonerFormatoEntero(text1(Index)) Then
                text2(Index).Text = PonerNombreDeCod(text1(Index), "scoope", "nomcoope")
                If text2(Index).Text = "" Then
                    cadMen = "No existe el Colectivo: " & text1(Index).Text & vbCrLf
                    cadMen = cadMen & "¿Desea crearlo?" & vbCrLf
                    If MsgBox(cadMen, vbQuestion + vbYesNo) = vbYes Then
                        Set frmCoo = New frmManCoope
                        frmCoo.DatosADevolverBusqueda = "0|1|"
                        frmCoo.NuevoCodigo = text1(Index).Text
                        text1(Index).Text = ""
                        TerminaBloquear
                        frmCoo.Show vbModal
                        Set frmCoo = Nothing
                        If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Data1, 1
                    Else
                        text1(Index).Text = ""
                    End If
                    PonerFoco text1(Index)
                End If
            Else
                text2(Index).Text = ""
            End If
            
        Case 8 'Situacion
            If PonerFormatoEntero(text1(Index)) Then
                text2(Index) = PonerNombreDeCod(text1(Index), "ssitua", "nomsitua", "codsitua", "N")
                If text2(Index).Text = "" Then
                    cadMen = "No existe la Situación: " & text1(Index).Text & vbCrLf
                    cadMen = cadMen & "¿Desea crearla?" & vbCrLf
                    If MsgBox(cadMen, vbQuestion + vbYesNo) = vbYes Then
                        Set frmSit = New frmManSitua
                        frmSit.DatosADevolverBusqueda = "0|1|"
                        frmSit.NuevoCodigo = text1(Index).Text
                        text1(Index).Text = ""
                        TerminaBloquear
                        frmSit.Show vbModal
                        Set frmSit = Nothing
                        If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Data1, 1
                    Else
                        text1(Index).Text = ""
                    End If
                    PonerFoco text1(Index)
                End If
            Else
                text2(Index).Text = ""
            End If
                
        Case 17 'Forma de Pago
            If PonerFormatoEntero(text1(Index)) Then
                text2(Index) = PonerNombreDeCod(text1(Index), "sforpa", "nomforpa", "codforpa", "N")
                If text2(Index).Text = "" Then
                    cadMen = "No existe la F.Pago: " & text1(Index).Text & vbCrLf
                    cadMen = cadMen & "¿Desea crearla?" & vbCrLf
                    If MsgBox(cadMen, vbQuestion + vbYesNo) = vbYes Then
                        Set frmFPa = New frmManFpago
                        frmFPa.DatosADevolverBusqueda = "0|1|"
                        frmFPa.NuevoCodigo = text1(Index).Text
                        text1(Index).Text = ""
                        TerminaBloquear
                        frmFPa.Show vbModal
                        Set frmFPa = Nothing
                        If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Data1, 1
                    Else
                        text1(Index).Text = ""
                    End If
                    PonerFoco text1(Index)
                End If
            Else
                text2(Index).Text = ""
            End If
        
        Case 15, 16 'Fechas
            PonerFormatoFecha text1(Index)
            
        Case 9 'cuenta contable
            If text1(Index).Text = "" Then Exit Sub
            If Modo = 3 Then
                text2(Index).Text = PonerNombreCuenta(text1(Index), Modo, "") 'text1(0).Text)
            Else
                text2(Index).Text = PonerNombreCuenta(text1(Index), Modo, text1(0).Text)
            End If
            
        Case 23 'DTO. X LITRO
            cadMen = TransformaPuntosComas(text1(Index).Text)
            text1(Index).Text = Format(cadMen, "0.0000")
            
        Case 25 'Nro de socio de la cooperativa
            PonerFormatoEntero text1(Index)
            
        Case 31 'Dia de Pago Fijo
            PonerFormatoEntero text1(Index)
            
            
        Case 32 ' codigo de iban
            text1(Index).Text = UCase(text1(Index).Text)
            
            
        '[Monica]02/07/2014: no formateabamos el banco y sucursal, estan en la base de datos como varchar
        Case 18 'banco
            If text1(Index).Text <> "" Then text1(Index).Text = Format(text1(Index).Text, "0000")
        Case 19 'sucursal
            If text1(Index).Text <> "" Then text1(Index).Text = Format(text1(Index).Text, "0000")
        
    End Select
    
    '[Monica]: calculo del iban si no lo ponen
    If Index = 18 Or Index = 19 Or Index = 20 Or Index = 21 Then
        Dim cta As String
        Dim CC As String
        
        If text1(18).Text <> "" And text1(19).Text <> "" And text1(20).Text <> "" And text1(21).Text <> "" Then
            
            cta = Format(text1(18).Text, "0000") & Format(text1(19).Text, "0000") & Format(text1(20).Text, "00") & Format(text1(21).Text, "0000000000")
            If Len(cta) = 20 Then
    '        Text1(42).Text = Calculo_CC_IBAN(cta, Text1(42).Text)
    
                If text1(32).Text = "" Then
                    'NO ha puesto IBAN
                    If DevuelveIBAN2("ES", cta, cta) Then text1(32).Text = "ES" & cta
                Else
                    CC = CStr(Mid(text1(32).Text, 1, 2))
                    If DevuelveIBAN2(CStr(CC), cta, cta) Then
                        If Mid(text1(32).Text, 3) <> cta Then
                            
                            MsgBox "Codigo IBAN distinto del calculado [" & CC & cta & "]", vbExclamation
                        End If
                    End If
                End If
                
            End If
        End If
    End If
    ' ***************************************************************************
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = teclaBuscar Then
        If Modo = 1 Or Modo = 3 Or Modo = 4 Then
            Select Case Index
                Case 7: KEYBusqueda KeyAscii, 0 'colectivo
                Case 8: KEYBusqueda KeyAscii, 1 'situacion
                Case 9: KEYBusqueda KeyAscii, 2 'cuenta contable
                Case 15: KEYFecha KeyAscii, 15 'fecha de alta
                Case 16: KEYFecha KeyAscii, 16 'fecha de baja
                Case 17: KEYBusqueda KeyAscii, 3 'forma pago
            End Select
        End If
    Else
        KEYpress KeyAscii
    End If
End Sub

Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
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

Private Sub KEYFecha(KeyAscii As Integer, indice As Integer)
    KeyAscii = 0
    imgFec_Click (indice)
End Sub


'************* LLINIES: ****************************
Private Sub ToolAux_ButtonClick(Index As Integer, ByVal Button As MSComctlLib.Button)
'-- pon el bloqueo aqui
    'If BLOQUEADesdeFormulario2(Me, Data1, 1) Then
    Select Case Button.Index
        Case 1
            BotonAnyadirLinea Index
        Case 2
            BotonModificarLinea Index
        Case 3
            BotonEliminarLinea Index
'        Case 5
'            BotonTarjetasLibres
        Case Else
    End Select
    'End If
End Sub

Private Sub BotonEliminarLinea(Index As Integer)
Dim SQL As String
Dim vWhere As String
Dim eliminar As Boolean

    On Error GoTo Error2

    ModoLineas = 3 'Posem Modo Eliminar Llínia
    
    If Modo = 4 Then 'Modificar Capçalera
        cmdAceptar_Click
        If ModoLineas = 0 Then Exit Sub
    End If
       
    NumTabMto = Index
    PonerModo 5, Index

    If AdoAux(Index).Recordset.EOF Then Exit Sub
    If Not SepuedeBorrar(Index) Then Exit Sub
    NumTabMto = Index
    eliminar = False
   
    vWhere = ObtenerWhereCab(True)
    
    ' ***** independentment de si tenen datagrid o no,
    ' canviar els noms, els formats i el DELETE *****
    Select Case Index
        Case 0 'tarjetas
            SQL = "¿Seguro que desea eliminar la Tarjeta?"
            SQL = SQL & vbCrLf & "Tarjeta: " & AdoAux(Index).Recordset!Numtarje
            If MsgBox(SQL, vbQuestion + vbYesNo) = vbYes Then
                eliminar = True
                SQL = "DELETE FROM starje"
                SQL = SQL & vWhere & " AND numlinea= " & AdoAux(Index).Recordset!NumLinea
            End If
            
        Case 1 'matriculas
            SQL = "¿Seguro que desea eliminar la Matricula?"
            SQL = SQL & vbCrLf & "Nombre: " & AdoAux(Index).Recordset!MATRICUL
            If MsgBox(SQL, vbQuestion + vbYesNo) = vbYes Then
                eliminar = True
                SQL = "DELETE FROM smatri"
                SQL = SQL & vWhere & " AND numlinea= " & AdoAux(Index).Recordset!NumLinea
            End If
            
            
        Case 2 'margenes
            SQL = "¿Seguro que desea eliminar el Margen del Artículo?"
            SQL = SQL & vbCrLf & "Código: " & AdoAux(Index).Recordset!codartic
            If MsgBox(SQL, vbQuestion + vbYesNo) = vbYes Then
                eliminar = True
                SQL = "DELETE FROM smargen"
                SQL = SQL & vWhere & " AND numlinea= " & AdoAux(Index).Recordset!NumLinea
            End If
        
            
    End Select

    If eliminar Then
        NumRegElim = AdoAux(Index).Recordset.AbsolutePosition
        TerminaBloquear
        Conn.Execute SQL
        ' *** si n'hi han tabs sense datagrid, posar l'If ***
        CargaGrid Index, True
        If Not SituarDataTrasEliminar(AdoAux(Index), NumRegElim, True) Then
'            PonerCampos
            
        End If
        If BLOQUEADesdeFormulario2(Me, Data1, 1) Then BotonModificar
        ' *** si n'hi han tabs ***
        SituarTab (NumTabMto + 1)
    End If
    
    ModoLineas = 0
    PosicionarData
    
    Exit Sub
Error2:
    Screen.MousePointer = vbDefault
    MuestraError Err.Number, "Eliminando linea", Err.Description
End Sub

Private Sub BotonAnyadirLinea(Index As Integer)
Dim NumF As String
Dim vWhere As String, vTabla As String
Dim anc As Single
Dim I As Integer
    
    If Index = 0 And vParamAplic.Cooperativa = 1 Then
        
        vWhere = ObtenerWhereCab(False)
        NumF = SugerirCodigoSiguienteStr("starje", "numlinea", vWhere)
        
        
        Set FrmCli2 = New frmManClien2
    
        FrmCli2.Socio = text1(0).Text
        FrmCli2.NuevoCodigo = text1(1).Text
        FrmCli2.NumLin = NumF
        FrmCli2.ModoExt = 3
                
        FrmCli2.Show vbModal
        
        PosicionarData
        
        Exit Sub
    End If
    
    
    ModoLineas = 1 'Posem Modo Afegir Llínia
    
    If (Modo = 3) Or (Modo = 4) Then 'Insertar o Modificar Capçalera
        cmdAceptar_Click
        If ModoLineas = 0 Then Exit Sub
    End If
       
    NumTabMto = Index
    PonerModo 5, Index
    
    ' *** bloquejar la clau primaria de la capçalera ***
    BloquearTxt text1(0), True

    ' *** posar el nom del les distintes taules de llínies ***
    Select Case Index
        Case 0: vTabla = "starje"
        Case 1: vTabla = "smatri"
        Case 2: vTabla = "smargen"
    End Select
    
    vWhere = ObtenerWhereCab(False)
    
    Select Case Index
        Case 0, 1, 2 ' *** pose els index dels tabs de llínies que tenen datagrid ***
            ' *** canviar la clau primaria de les llínies,
            'pasar a "" si no volem que mos sugerixca res a l'afegir ***
            NumF = SugerirCodigoSiguienteStr(vTabla, "numlinea", vWhere)

            AnyadirLinea DataGridAux(Index), AdoAux(Index)
    
            anc = DataGridAux(Index).Top
            If DataGridAux(Index).Row < 0 Then
                anc = anc + 210
            Else
                anc = anc + DataGridAux(Index).RowTop(DataGridAux(Index).Row) + 5
            End If
            
            LLamaLineas Index, ModoLineas, anc
        
            Select Case Index
                ' *** valor per defecte a l'insertar i formateig de tots els camps ***
                Case 0 'tarjetas
                    txtAux(0).Text = text1(0).Text 'codsocio
                    txtAux(3).Text = text1(1).Text 'nomsocio
                    txtAux(1).Text = NumF 'numlinea
                    txtAux(2).Text = NumTarj
                    For I = 4 To 7
                        txtAux(I).Text = ""
                    Next I
                    txtAux(12).Text = ""
                    cmbAux(0).ListIndex = 1
                    PonerFoco txtAux(2)
                Case 1 'matriculas
                    txtAux(8).Text = text1(0).Text 'codsocio
                    txtAux(9).Text = NumF 'numlinea
                    For I = 10 To 11
                        txtAux(I).Text = ""
                    Next I
                    
                    PonerFoco txtAux(10)
                Case 2 'margenes
                    txtAux(13).Text = text1(0).Text 'codsocio
                    txtAux(14).Text = NumF 'numlinea
                    For I = 15 To 17
                        txtAux(I).Text = ""
                    Next I
                    text2(0).Text = ""
                    PonerFoco txtAux(15)
            End Select
    End Select
End Sub

Private Sub BotonModificarLinea(Index As Integer)
    Dim anc As Single
    Dim I As Integer
    Dim J As Integer
    
    If AdoAux(Index).Recordset.EOF Then Exit Sub
    If AdoAux(Index).Recordset.RecordCount < 1 Then Exit Sub
    
    If Index = 0 And vParamAplic.Cooperativa = 1 Then
        
        
        Set FrmCli2 = New frmManClien2
    
        FrmCli2.Socio = text1(0).Text
        FrmCli2.NuevoCodigo = text1(1).Text
        FrmCli2.NumLin = AdoAux(0).Recordset!NumLinea
        FrmCli2.ModoExt = 4
                
        FrmCli2.Show vbModal
        
        PosicionarData
        
        Exit Sub
    End If
    
    
    
    
    
    
    
    ModoLineas = 2 'Modificar llínia
       
    If Modo = 4 Then 'Modificar Capçalera
        cmdAceptar_Click
        If ModoLineas = 0 Then Exit Sub
    End If
       
    NumTabMto = Index
    PonerModo 5, Index
    ' *** bloqueje la clau primaria de la capçalera ***
    BloquearTxt text1(0), True
  
    Select Case Index
        Case 0, 1, 2 ' *** pose els index de llínies que tenen datagrid (en o sense tab) ***
            If DataGridAux(Index).Bookmark < DataGridAux(Index).FirstRow Or DataGridAux(Index).Bookmark > (DataGridAux(Index).FirstRow + DataGridAux(Index).VisibleRows - 1) Then
                I = DataGridAux(Index).Bookmark - DataGridAux(Index).FirstRow
                DataGridAux(Index).Scroll 0, I
                DataGridAux(Index).Refresh
            End If
              
            anc = DataGridAux(Index).Top
            If DataGridAux(Index).Row < 0 Then
                anc = anc + 210
            Else
                anc = anc + DataGridAux(Index).RowTop(DataGridAux(Index).Row) + 5
            End If
    End Select
    
    Select Case Index
        ' *** valor per defecte al modificar dels camps del grid ***
        Case 0 'TARJETAS
        
            For J = 0 To 2
                txtAux(J).Text = DataGridAux(Index).Columns(J).Text
            Next J
            
            PosicionarCombo cmbAux(0), AdoAux(Index).Recordset!tiptarje
    
            For J = 3 To 3
                txtAux(J).Text = DataGridAux(Index).Columns(J + 2).Text
            Next J
            
            '[Monica]22/11/2013: iban
            txtAux(18).Text = DataGridAux(Index).Columns(6).Text
            
            For J = 4 To 7
                txtAux(J).Text = DataGridAux(Index).Columns(J + 3).Text
            Next J
            txtAux(12).Text = DataGridAux(Index).Columns(11).Text
            For I = 0 To 1
                BloquearTxt txtAux(I), False
            Next I
            
        Case 1 'MATRICULAS
            For J = 8 To 11
                txtAux(J).Text = DataGridAux(Index).Columns(J - 8).Text
            Next J
            
            For I = 8 To 9
                BloquearTxt txtAux(I), False
            Next I
            
        Case 2 'MARGENES
        
            For J = 13 To 14
                txtAux(J).Text = DataGridAux(Index).Columns(J - 13).Text
            Next J
            txtAux(15).Text = DataGridAux(Index).Columns(2).Text
            text2(0).Text = DataGridAux(Index).Columns(3).Text
            txtAux(16).Text = DataGridAux(Index).Columns(4).Text
            
            For I = 13 To 14
                BloquearTxt txtAux(I), False
            Next I
    End Select
    
    LLamaLineas Index, ModoLineas, anc
   
    ' *** foco al 1r camp visible de les llinies en grids que no siga PK (en o sense tab) ***
    Select Case Index
        Case 0 'TARJETAS
            PonerFoco txtAux(2)
        Case 1 'MATRICULAS
            PonerFoco txtAux(10)
        Case 2 'MARGENES
            PonerFoco txtAux(16)
    End Select
    ' ***************************************************************************************
End Sub

Private Sub BotonClientesLibres()

    Dim frmMen As frmMensaje
    Set frmMen = New frmMensaje
    frmMen.OpcionMensaje = 5
    frmMen.Show vbModal
    If frmMen.pTitulo <> "" Then
        NumTarj = frmMen.pTitulo
        txtAux(2).Text = NumTarj
    End If
    Set frmMen = Nothing
End Sub

Private Sub BotonTarjetasLibres()

'    frmMensaje.OpcionMensaje = 4
'    frmMensaje.Show vbModal

    Dim frmMen As frmMensaje
    Set frmMen = New frmMensaje
    frmMen.OpcionMensaje = 4
    frmMen.Show vbModal
    If frmMen.pTitulo <> "" Then
        NumTarj = frmMen.pTitulo
        txtAux(2).Text = NumTarj
    End If
    Set frmMen = Nothing
End Sub

Private Sub LLamaLineas(Index As Integer, xModo As Byte, Optional alto As Single)
Dim jj As Integer
Dim b As Boolean

    ' *** si n'hi han tabs sense datagrid posar el If ***
    DeseleccionaGrid DataGridAux(Index)
       
    b = (xModo = 1 Or xModo = 2) 'Insertar o Modificar Llínies
    Select Case Index
        Case 0 'tarjetas
             For jj = 2 To 7
                txtAux(jj).visible = b
                txtAux(jj).Top = alto
            Next jj
            txtAux(12).visible = b
            txtAux(12).Top = alto
            txtAux(18).visible = b
            txtAux(18).Top = alto
            
            cmbAux(0).visible = b
            cmbAux(0).Top = alto - 15
            
        Case 1 'matriculas
            For jj = 10 To 11
                txtAux(jj).visible = b
                txtAux(jj).Top = alto
            Next jj
            
        Case 2 ' MARGENES
            For jj = 16 To 17
                txtAux(jj).visible = b
                txtAux(jj).Top = alto
            Next jj
            text2(0).visible = (xModo = 1)
            text2(0).Top = alto
            txtAux(15).visible = (xModo = 1)
            txtAux(15).Top = alto
            Me.btnBuscar(0).visible = (xModo = 1)
            Me.btnBuscar(0).Enabled = (xModo = 1)
            Me.btnBuscar(0).Top = alto
            
    End Select
End Sub

' ********* si n'hi han combos a la capçalera ************
Private Sub CargaCombo(Index As Integer)
Dim Ini As Integer
Dim Fin As Integer
Dim I As Integer

    ' *** neteje els combos, els pose valor i seleccione el valor per defecte ***
    For I = 0 To 1
        Combo1(I).Clear
    Next I
    cmbAux(0).Clear
    
'    For i = 2 To 3
'        Combo1(i).AddItem "No"
'        Combo1(i).ItemData(Combo1(i).NewIndex) = 0
'        Combo1(i).AddItem "Si"
'        Combo1(i).ItemData(Combo1(i).NewIndex) = 1
'    Next i
'
    Combo1(0).AddItem "Comercio"
    Combo1(0).ItemData(Combo1(0).NewIndex) = 2
    Combo1(0).AddItem "Particular"
    Combo1(0).ItemData(Combo1(0).NewIndex) = 0
    Combo1(0).AddItem "Profesional"
    Combo1(0).ItemData(Combo1(0).NewIndex) = 1
    
    cmbAux(0).AddItem "Bonificado"
    cmbAux(0).ItemData(cmbAux(0).NewIndex) = 1
    cmbAux(0).AddItem "Normal"
    cmbAux(0).ItemData(cmbAux(0).NewIndex) = 0
    cmbAux(0).AddItem "Profesional"
    cmbAux(0).ItemData(cmbAux(0).NewIndex) = 2

    ' combo para el listado de estadisticas de articulos
    Combo1(1).AddItem "Cooperativa"
    Combo1(1).ItemData(Combo1(1).NewIndex) = 0
    Combo1(1).AddItem "Tarjetas Visa"
    Combo1(1).ItemData(Combo1(1).NewIndex) = 1
    Combo1(1).AddItem "Crédito local"
    Combo1(1).ItemData(Combo1(1).NewIndex) = 2
    Combo1(1).AddItem "Clientes paso"
    Combo1(1).ItemData(Combo1(1).NewIndex) = 3
    Combo1(1).AddItem "Efectivo"
    Combo1(1).ItemData(Combo1(1).NewIndex) = 4


End Sub

Private Sub txtAux_LostFocus(Index As Integer)
Dim cadMen As String
Dim Nuevo As Boolean
Dim Cant As Long
    If Not PerderFocoGnral(txtAux(Index), Modo) Then Exit Sub
    
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub
    

    ' ******* configurar el LostFocus dels camps de llínies (dins i fora grid) ********
    Select Case Index
        Case 3, 10, 12 ' Nomtarje y matricula
            txtAux(Index).Text = UCase(txtAux(Index).Text)
            
        Case 4, 5 ' banco y sucursal con 4 digitos
            txtAux(Index).Text = Format(txtAux(Index).Text, "0000")
            
        Case 12, 11 'Observaciones dpto
            PonerFocoBtn Me.cmdAceptar
            
        Case 15 ' articulo
            If PonerFormatoEntero(txtAux(Index)) Then
                text2(0).Text = PonerNombreDeCod(txtAux(Index), "sartic", "nomartic", "codartic", "N")
                If text2(0).Text = "" Then
                    cadMen = "No existe el Articulo: " & txtAux(Index).Text & vbCrLf
                    MsgBox cadMen, vbExclamation
                Else
                    Cant = TotalRegistros("select count(*) from sartic inner join sfamia on sartic.codfamia = sfamia.codfamia where sartic.codartic = " & DBSet(txtAux(15).Text, "N") & " and sfamia.tipfamia = 1")
                    If Cant = 0 Then
                        MsgBox "El artículo introducido no es de la familia de combustibles. Reintroduzca.", vbExclamation
                        PonerFoco txtAux(15)
                    End If
                End If
            End If
        
        Case 16 ' porcentaje de margen
            PonerFormatoDecimal txtAux(Index), 4
        
        Case 17 ' euros por litro
            PonerFormatoDecimal txtAux(Index), 7
            PonerFocoBtn Me.cmdAceptar
            
        Case 18 ' codigo de iban
            txtAux(Index).Text = UCase(txtAux(Index).Text)
            
    End Select
    
    '[Monica]: calculo del iban si no lo ponen
    If Index = 4 Or Index = 5 Or Index = 6 Or Index = 7 Then
        Dim cta As String
        Dim CC As String
        If txtAux(4).Text <> "" And txtAux(5).Text <> "" And txtAux(6).Text <> "" And txtAux(7).Text <> "" Then
            
            cta = Format(txtAux(4).Text, "0000") & Format(txtAux(5).Text, "0000") & Format(txtAux(6).Text, "00") & Format(txtAux(7).Text, "0000000000")
            If Len(cta) = 20 Then
    '        Text1(42).Text = Calculo_CC_IBAN(cta, Text1(42).Text)
    
                If txtAux(18).Text = "" Then
                    'NO ha puesto IBAN
                    If DevuelveIBAN2("ES", cta, cta) Then txtAux(18).Text = "ES" & cta
                Else
                    CC = CStr(Mid(txtAux(18).Text, 1, 2))
                    If DevuelveIBAN2(CStr(CC), cta, cta) Then
                        If Mid(txtAux(18).Text, 3) <> cta Then
                            
                            MsgBox "Codigo IBAN distinto del calculado [" & CC & cta & "]", vbExclamation
                        End If
                    End If
                End If
                
            End If
        End If
    End If
            
End Sub

Private Sub txtAux_GotFocus(Index As Integer)
   If Not txtAux(Index).MultiLine Then ConseguirFocoLin txtAux(Index)
End Sub


Private Sub txtAux_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If Not txtAux(Index).MultiLine Then KEYdown KeyCode
End Sub

Private Sub txtAux_KeyPress(Index As Integer, KeyAscii As Integer)
    If Not txtAux(Index).MultiLine Then
        If KeyAscii = teclaBuscar Then
            If Modo = 5 And (ModoLineas = 1 Or ModoLineas = 2) Then
            End If
        Else
            KEYpress KeyAscii
        End If
    End If
End Sub

Private Function DatosOkLlin(nomframe As String) As Boolean
Dim RS As ADODB.Recordset
Dim SQL As String
Dim b As Boolean
Dim Cant As Integer
Dim Mens As String
Dim vFact As Byte, vDocum As Byte
Dim cta As String
Dim cadMen As String
    On Error GoTo EDatosOKLlin

    Mens = ""
    DatosOkLlin = False
        
    b = CompForm2(Me, 2, nomframe) 'Comprovar formato datos ok
    If Not b Then Exit Function
    
    Select Case NumTabMto
        Case 0 ' tarjetas
            ' en el caso de que la tarjeta sea profesional la matricula es obligatoria
            If cmbAux(0).ListIndex = 2 And txtAux(12).Text = "" Then
                MsgBox "Si el tipo de tarjeta es Profesional, debe introducir la matrícula", vbExclamation
                PonerFoco txtAux(12)
                b = False
            End If
            If b Then
            
                '[Monica]22/11/2013: añadida la comprobacion de que la cuenta contable sea correcta
                If txtAux(4).Text = "" Or txtAux(5).Text = "" Or txtAux(6).Text = "" Or txtAux(7).Text = "" Then
                    '[Monica]20/11/2013: añadido el codigo de iban
                    txtAux(18).Text = ""
                    txtAux(4).Text = ""
                    txtAux(5).Text = ""
                    txtAux(6).Text = ""
                    txtAux(7).Text = ""
                Else
                    cta = Format(txtAux(4).Text, "0000") & Format(txtAux(5).Text, "0000") & Format(txtAux(6).Text, "00") & Format(txtAux(7).Text, "0000000000")
                    If Val(ComprobarCero(cta)) = 0 Then
                        cadMen = "El cliente no tiene asignada cuenta bancaria."
                        MsgBox cadMen, vbExclamation
                    End If
                    If Not Comprueba_CC(cta) Then
                        cadMen = "La cuenta bancaria del cliente no es correcta. ¿ Desea continuar ?."
                        If MsgBox(cadMen, vbQuestion + vbYesNo + vbDefaultButton1) = vbYes Then
                            b = True
                        Else
                            PonerFoco txtAux(18)
                            b = False
                        End If
                    Else
        '                '[Monica]20/11/2013: añadimos el tema de la comprobacion del IBAN
        '                If Not Comprueba_CC_IBAN(cta, Text1(42).Text) Then
        '                    cadMen = "La cuenta IBAN del cliente no es correcta. ¿ Desea continuar ?."
        '                    If MsgBox(cadMen, vbQuestion + vbYesNo + vbDefaultButton1) = vbYes Then
        '                        b = True
        '                    Else
        '                        PonerFoco Text1(42)
        '                        b = False
        '                    End If
        '                End If
        
        '       sustituido por lo de David
                        BuscaChekc = ""
                        If Me.txtAux(18).Text <> "" Then BuscaChekc = Mid(txtAux(18).Text, 1, 2)
                            
                        If DevuelveIBAN2(BuscaChekc, cta, cta) Then
                            If Me.txtAux(18).Text = "" Then
                                If MsgBox("Poner IBAN ?", vbQuestion + vbYesNo) = vbYes Then Me.txtAux(18).Text = BuscaChekc & cta
                            Else
                                If Mid(txtAux(18).Text, 3) <> cta Then
                                    cta = "Calculado : " & BuscaChekc & cta
                                    cta = "Introducido: " & Me.txtAux(18).Text & vbCrLf & cta & vbCrLf
                                    cta = "Error en codigo IBAN" & vbCrLf & cta & "Continuar?"
                                    If MsgBox(cta, vbQuestion + vbYesNo) = vbNo Then
                                        PonerFoco txtAux(18)
                                        b = False
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            
            
            
            
            End If
        
        Case 2 ' margenes
            ' el codigo de articulos de los que vamos a introducir margenes tienen que ser necesariamente de la familia de combustibles
            If b Then
                Cant = TotalRegistros("select count(*) from sartic inner join sfamia on sartic.codfamia = sfamia.codfamia where sartic.codartic = " & DBSet(txtAux(15).Text, "N") & " and sfamia.tipfamia = 1")
                If Cant = 0 Then
                    MsgBox "El artículo introducido no es de la familia de combustibles. Reintroduzca.", vbExclamation
                    PonerFoco txtAux(15)
                    b = False
                End If
            End If
            ' comprobamos que el codigo de articulo no es haya introducido ya para el cliente
            If b And ModoLineas = 1 Then
                Cant = TotalRegistros("select count(*) from smargen where codsocio = " & DBSet(txtAux(13).Text, "N") & " and codartic = " & DBSet(txtAux(15).Text, "N"))
                If Cant <> 0 Then
                    MsgBox "Este artículo ya ha sido introducido. Modifique.", vbExclamation
                    PonerFoco txtAux(15)
                    b = False
                End If
            End If
            
            '[Monica]15/12/2011: No permitimos que columna de % de margen y euros/kilo esten en blanco
            If b And (ModoLineas = 1 Or ModoLineas = 2) Then
                If ComprobarCero(txtAux(16).Text) = 0 And ComprobarCero(txtAux(17).Text) = 0 Then
                    MsgBox "Alguna de las dos columnas % Margen o €/litro ha de tener un valor. Reintroduzca.", vbExclamation
                    PonerFoco txtAux(16)
                    b = False
                Else
                    If ComprobarCero(txtAux(16).Text) <> 0 And ComprobarCero(txtAux(17).Text) <> 0 Then
                        MsgBox "Las columnas % Margen y €/litro no pueden ser a la vez distinto de 0. Reintroduzca.", vbExclamation
                        PonerFoco txtAux(16)
                        b = False
                    End If
                End If
            End If
            
    End Select
    
    ' *** si cal fer atres comprovacions a les llínies (en o sense tab) ***
'    Select Case NumTabMto
'        Case 0  'CUENTAS BANCARIAS
'            SQL = "SELECT COUNT(ctaprpal) FROM cltebanc "
'            SQL = SQL & ObtenerWhereCab(True) & " AND ctaprpal=1"
'            If ModoLineas = 2 Then SQL = SQL & " AND numlinea<> " & AdoAux(NumTabMto).Recordset!NumLinea
'            Set RS = New ADODB.Recordset
'            RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
'            Cant = IIf(Not RS.EOF, RS.Fields(0).Value, 0)
'
'            RS.Close
'            Set RS = Nothing
'
'            'no n'hi ha cap conter principal i ha seleccionat que no
'            If (Cant = 0) And (cmbAux(10).ItemData(cmbAux(10).ListIndex) = 0) Then
'                Mens = "Debe una haber una cuenta principal"
'            ElseIf (cmbAux(10).ItemData(cmbAux(10).ListIndex) = 1) And (cmbAux(9).ItemData(cmbAux(9).ListIndex) = 0) Then
'                Mens = "Debe seleccionar que esta cuenta está activa si desea que sea la principal"
'            End If
'
''            'No puede haber más de una cuenta principal
''            If cant > 0 And (cmbAux(10).ItemData(cmbAux(10).ListIndex) = 1) Then
''                Mens = "No puede haber más de una cuenta principal."
''            End If
'
'            'No pueden haber registros con el mismo: codbanco-codsucur-digcontr-ctabanc
'            If Mens = "" Then
'                SQL = "SELECT count(codclien) FROM cltebanc "
'                SQL = SQL & " WHERE codclien=" & text1(0).Text & " AND codempre= " & vSesion.Empresa
'                If ModoLineas = 2 Then SQL = SQL & " AND numlinea<> " & AdoAux(NumTabMto).Recordset!NumLinea
'                SQL = SQL & " AND codnacio=" & cmbAux(0).ItemData(cmbAux(0).ListIndex)
'                SQL = SQL & " AND codbanco=" & txtAux(3).Text & " AND codsucur=" & txtAux(4).Text
'                SQL = SQL & " AND digcontr='" & txtAux(5).Text & "' AND ctabanco='" & txtAux(6).Text & "'"
'                Set RS = New ADODB.Recordset
'                RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
'                Cant = IIf(Not RS.EOF, RS.Fields(0).Value, 0)
'                If Cant > 0 Then
'                    Mens = "Ya Existe la cuenta bancaria: " & cmbAux(0).List(cmbAux(0).ListIndex) & " - " & txtAux(3).Text & "-" & txtAux(4).Text & "-" & txtAux(5).Text & "-" & txtAux(6).Text
'                End If
'                RS.Close
'                Set RS = Nothing
'            End If
'
'            If Mens <> "" Then
'                Screen.MousePointer = vbNormal
'                MsgBox Mens, vbExclamation
'                DatosOkLlin = False
'                'PonerFoco txtAux(3)
'                Exit Function
'            End If
'
'        Case 1 'DEPARTAMENTOS
'            'Solo puede haber un departamento principal de facturacion y uno de documentos
'            'o uno de ambas cosas
'            vFact = Me.cmbAux(18).ItemData(cmbAux(18).ListIndex)
'            vDocum = Me.cmbAux(19).ItemData(cmbAux(19).ListIndex)
'            'si he marcado que va a ser dpto principal comprobar
'            If Me.cmbAux(20).ItemData(cmbAux(20).ListIndex) = 1 Then
'                If (vFact = 0 And vDocum = 0) Then
'                'facturar o document o ambos debe ser 1
'                    Mens = "Debe indicar si va a ser el Departamento principal para: " & vbCrLf
'                    Mens = Mens & "   - Facturación" & vbCrLf & "   - Documentación" & vbCrLf & "   - Ambos"
'                    MsgBox Mens, vbInformation
'                    DatosOkLlin = False
'                    Exit Function
'                Else
'                    SQL = "SELECT count(princpal) FROM cltedpto "
'                    SQL = SQL & " WHERE codclien=" & text1(0).Text & " AND codempre= " & vSesion.Empresa
'                    If ModoLineas = 2 Then SQL = SQL & " AND numlinea<> " & AdoAux(NumTabMto).Recordset!NumLinea
''                    vWhere = " WHERE codprove=" & text1(0).Text & " AND codempre= " & codEmpre
'                    If vFact = 1 And vDocum = 0 Then
'                        SQL = SQL & " AND princpal=1 AND facturac=1 "
'                        Mens = "Ya existe un Departamento Principal de Facturación."
'                    ElseIf vFact = 0 And vDocum = 1 Then
'                        SQL = SQL & " AND princpal=1 AND document=1 "
'                        Mens = "Ya existe un Departamento Principal de Documentación."
'                    Else
'                        SQL = SQL & " AND princpal=1 AND (facturac=1 or document=1) "
'                        Mens = "Ya existe un Departamento Principal para Facturación o Documentación."
'                    End If
'
'                    Set RS = New ADODB.Recordset
'                    RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
'                    Cant = IIf(Not RS.EOF, RS.Fields(0).Value, 0)
'                    If Cant > 0 Then
'                        MsgBox Mens, vbExclamation
'                        DatosOkLlin = False
'                        Exit Function
'                    End If
'                    RS.Close
'                    Set RS = Nothing
'                End If
'            Else
'                'Ver si hay ya algun dpto principal, si es el primero que insertamos
'                'tiene que ser el principal
'                SQL = "SELECT COUNT(princpal) FROM cltedpto "
'                SQL = SQL & " WHERE codclien=" & text1(0).Text & " AND codempre= " & vSesion.Empresa & " AND princpal=1"
'
'                Set RS = New ADODB.Recordset
'                RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
'                Cant = IIf(Not RS.EOF, RS.Fields(0).Value, 0)
'
'                RS.Close
'                Set RS = Nothing
'
'                'no n'hi ha cap conter principal i ha seleccionat que no
'                If (Cant = 0) Then
'                    Mens = "Debe haber una Departamento principal"
''                ElseIf (cmbAux(10).ItemData(cmbAux(10).ListIndex) = 1) And (cmbAux(9).ItemData(cmbAux(9).ListIndex) = 0) Then
''                    Mens = "Debe seleccionar que esta cuenta está activa si desea que sea la principal"
'                End If
'
'                If Mens <> "" Then
'                   MsgBox Mens, vbExclamation
'                   DatosOkLlin = False
'                   Exit Function
'                End If
'
'            End If
'
'        Case 4 'comisiones
'            If b And (ModoLineas = 1) Then 'insertar llínia
'                'Datos = DevuelveDesdeBDnew(cPTours, "pobrecog", "lugrecog", "lugrecog", txtAux(1).Text, "N", , "codpobla", txtAux(0).Text, "N")
'
'                SQL = "SELECT DISTINCT COUNT(codempre) FROM cltcomis WHERE codempre = " & vSesion.Empresa
'                SQL = SQL & " AND codclien = " & txtAux(53).Text
'                SQL = SQL & " AND codprodu = " & txtAux(54).Text
'
'                Set RS = New ADODB.Recordset
'                RS.Open SQL, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
'
'                Cant = IIf(Not RS.EOF, RS.Fields(0).Value, 0)
'
'                RS.Close
'                Set RS = Nothing
'
'                 If Cant > 0 Then
'                    MsgBox "Ya existe la Comisión del Producto: " & txtAux(54).Text & " para este Cliente", vbExclamation
'                    DatosOkLlin = False
'                    PonerFoco txtAux(54) '*** posar el foco al 1r camp visible de la PK de les llínies ***
'                    Exit Function
'                 End If
'            End If
'    End Select
    ' ******************************************************************************
    DatosOkLlin = b
    
EDatosOKLlin:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Function

Private Function SepuedeBorrar(ByRef Index As Integer) As Boolean
    SepuedeBorrar = False
    
    ' *** si cal comprovar alguna cosa abans de borrar ***
'    Select Case Index
'        Case 0 'cuentas bancarias
'            If AdoAux(Index).Recordset!ctaprpal = 1 Then
'                MsgBox "No puede borrar una Cuenta Principal. Seleccione antes otra cuenta como Principal", vbExclamation
'                Exit Function
'            End If
'    End Select
    ' ****************************************************
    
    SepuedeBorrar = True
End Function

Private Sub imgBuscar_Click(Index As Integer)
    TerminaBloquear
     Select Case Index
        Case 0 'colectivos
            Set frmCoo = New frmManCoope
            frmCoo.DatosADevolverBusqueda = "0|1|"
            frmCoo.CodigoActual = text1(7).Text
            frmCoo.Show vbModal
            Set frmCoo = Nothing
            PonerFoco text1(7)
            
        Case 1 'situaciones
            Set frmSit = New frmManSitua
            frmSit.DatosADevolverBusqueda = "0|1|"
            frmSit.CodigoActual = text1(8).Text
            frmSit.Show vbModal
            Set frmSit = Nothing
            PonerFoco text1(8)
            
        Case 3 'formas de pago
            Set frmFPa = New frmManFpago
            frmFPa.DatosADevolverBusqueda = "0|1|"
            frmFPa.CodigoActual = text1(17).Text
            frmFPa.Show vbModal
            Set frmFPa = Nothing
            PonerFoco text1(17)
            
        Case 2 'Cuentas Contables (de contabilidad)
            If vParamAplic.NumeroConta = 0 Then Exit Sub
            
            indice = Index + 7
            Set frmCtas = New frmCtasConta
            frmCtas.NumDigit = 0
            frmCtas.DatosADevolverBusqueda = "0|1|"
            frmCtas.CodigoActual = txtAux(indice).Text
            frmCtas.Show vbModal
            Set frmCtas = Nothing
            PonerFoco text1(indice)
            
    End Select
    
    If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Data1, 1
End Sub

Private Sub frmCtas_DatoSeleccionado(CadenaSeleccion As String)
'Cuentas contables de la Contabilidad
    text1(indice).Text = RecuperaValor(CadenaSeleccion, 1) 'codmacta
    text2(indice).Text = RecuperaValor(CadenaSeleccion, 2) 'des macta
End Sub

Private Sub frmCoo_DatoSeleccionado(CadenaSeleccion As String)
'Mantenimiento de Colectivos
    text1(7).Text = RecuperaValor(CadenaSeleccion, 1) 'codcoope
    FormateaCampo text1(7)
    text2(7).Text = RecuperaValor(CadenaSeleccion, 2) 'nomcoope
End Sub

Private Sub frmSit_DatoSeleccionado(CadenaSeleccion As String)
'Mantenimiento Grupo
    text1(8).Text = RecuperaValor(CadenaSeleccion, 1) 'codsitua
    FormateaCampo text1(8)
    text2(8).Text = RecuperaValor(CadenaSeleccion, 2) 'nomsitua
End Sub

Private Sub frmFpa_DatoSeleccionado(CadenaSeleccion As String)
'Mantenimiento F.Pago
    text1(17).Text = RecuperaValor(CadenaSeleccion, 1) 'codforpa
    FormateaCampo text1(17)
    text2(17).Text = RecuperaValor(CadenaSeleccion, 2) 'nomforpa
End Sub


Private Sub imgWeb_Click(Index As Integer)
    'Abrimos el explorador de windows con la pagina Web del cliente
    If Modo = 0 Then Exit Sub
    Screen.MousePointer = vbHourglass
'    If LanzaHome("websoporte") Then espera 2
    If LanzaHomeGnral(text1(10).Text) Then espera 2
    Screen.MousePointer = vbDefault
End Sub

Private Sub DataGridAux_RowColChange(Index As Integer, LastRow As Variant, ByVal LastCol As Integer)
Dim I As Byte

    If ModoLineas <> 1 Then
        Select Case Index
            Case 0 'tarjetas
                If DataGridAux(Index).Columns.Count > 2 Then
'                    txtAux(11).Text = DataGridAux(Index).Columns("direccio").Text
'                    txtAux(12).Text = DataGridAux(Index).Columns("observac").Text
                End If
                
            Case 1 'matriculas
                If DataGridAux(Index).Columns.Count > 2 Then
'                    txtAux(21).Text = DataGridAux(Index).Columns(5).Text
'                    txtAux(22).Text = DataGridAux(Index).Columns(6).Text
'                    txtAux(23).Text = DataGridAux(Index).Columns(8).Text
'                    txtAux(24).Text = DataGridAux(Index).Columns(15).Text
'                    txtAux2(22).Text = DataGridAux(Index).Columns(7).Text
                End If
                
        End Select
        
    Else 'vamos a Insertar
        Select Case Index
            Case 0 'tarjetas
'                txtAux(11).Text = ""
'                txtAux(12).Text = ""
            Case 1 'matriculas
                For I = 21 To 24
'                   txtAux(i).Text = ""
                Next I
'               txtAux2(22).Text = ""
            Case 2 'Margenes
'               txtAux(50).Text = ""
'               txtAux(51).Text = ""
        End Select
    End If
End Sub

' ***** si n'hi han varios nivells de tabs *****
Private Sub SituarTab(numTab As Integer)
    On Error Resume Next
    
    SSTab1.Tab = numTab
    
    If Err.Number <> 0 Then Err.Clear
End Sub
' **********************************************

Private Sub CargaFrame(Index As Integer, enlaza As Boolean)
Dim tip As Integer
Dim I As Byte

    AdoAux(Index).ConnectionString = Conn
    AdoAux(Index).RecordSource = MontaSQLCarga(Index, enlaza)
    AdoAux(Index).CursorType = adOpenDynamic
    AdoAux(Index).LockType = adLockPessimistic
    AdoAux(Index).Refresh
    
    If Not AdoAux(Index).Recordset.EOF Then
        PonerCamposForma2 Me, AdoAux(Index), 2, "FrameAux" & Index
    Else
        ' *** si n'hi han tabs sense datagrids, li pose els valors als camps ***
        NetejaFrameAux "FrameAux3" 'neteja només lo que te TAG
    End If
End Sub

' *** si n'hi han tabs sense datagrids ***
Private Sub NetejaFrameAux(nom_frame As String)
Dim Control As Object
    
    For Each Control In Me.Controls
        If (Control.Tag <> "") Then
            If (Control.Container.Name = nom_frame) Then
                If TypeOf Control Is TextBox Then
                    Control.Text = ""
                ElseIf TypeOf Control Is ComboBox Then
                    Control.ListIndex = -1
                End If
            End If
        End If
    Next Control

End Sub

Private Sub CargaGrid(Index As Integer, enlaza As Boolean)
Dim b As Boolean
Dim I As Byte
Dim tots As String

    On Error GoTo ECarga

    tots = MontaSQLCarga(Index, enlaza)

    CargaGridGnral Me.DataGridAux(Index), Me.AdoAux(Index), tots, PrimeraVez
    
    Select Case Index
        Case 0 'tarjetas
            'si es visible|control|tipo campo|nombre campo|ancho control|
            tots = "N||||0|;N||||0|;" 'codsocio,numlinea
            tots = tots & "S|txtAux(2)|T|Tarjeta|1000|;N||||0|;S|cmbAux(0)|C|Tipo|1000|;"
            tots = tots & "S|txtAux(3)|T|Nombre|3300|;S|txtAux(18)|T|IBAN|650|;S|txtAux(4)|T|Banco|650|;"
            tots = tots & "S|txtAux(5)|T|Sucur.|650|;S|txtAux(6)|T|DC|380|;"
            tots = tots & "S|txtAux(7)|T|Cuenta|1500|;"
            tots = tots & "S|txtAux(12)|T|Matricula|1200|;"
            
            arregla tots, DataGridAux(Index), Me
        
'            DataGridAux(0).Columns(5).Alignment = dbgRight
'            DataGridAux(0).Columns(6).Alignment = dbgRight
'            DataGridAux(0).Columns(7).Alignment = dbgRight
'            DataGridAux(0).Columns(8).Alignment = dbgRight
        
            b = (Modo = 4) And ((ModoLineas = 1) Or (ModoLineas = 2))
            
        Case 1 'Matriculas
            'si es visible|control|tipo campo|nombre campo|ancho control|
            tots = "N||||0|;N||||0|;" 'codsocio,numlinea
            tots = tots & "S|txtAux(10)|T|Matricula|1700|;"
            tots = tots & "S|txtAux(11)|T|Observacion|3000|;"

            arregla tots, DataGridAux(Index), Me
            
            b = (Modo = 4) And ((ModoLineas = 1) Or (ModoLineas = 2))
            
        Case 2 'Margenes de precios de combustibles
            'si es visible|control|tipo campo|nombre campo|ancho control|
            tots = "N||||0|;N||||0|;" 'codsocio,numlinea
            tots = tots & "S|txtAux(15)|T|Artículo|1000|;S|btnBuscar(0)|B|||;S|text2(0)|T|Nombre|5000|;"
            tots = tots & "S|txtAux(16)|T|% Margen|1400|;"
            tots = tots & "S|txtAux(17)|T|Margen €/Litro|1400|;"

            arregla tots, DataGridAux(Index), Me
            
            b = (Modo = 4) And ((ModoLineas = 1) Or (ModoLineas = 2))
            
            
    End Select
    
    DataGridAux(Index).ScrollBars = dbgAutomatic
      
    ' **** si n'hi han llínies en grids i camps fora d'estos ****
    If Not AdoAux(Index).Recordset.EOF Then
        DataGridAux_RowColChange Index, 1, 1
    Else
'        LimpiarCamposFrame Index
    End If
      
ECarga:
    If Err.Number <> 0 Then MuestraError Err.Number, "Cargando datos grid: " & DataGridAux(Index).Tag, Err.Description
End Sub

Private Sub InsertarLinea()
'Inserta registre en les taules de Llínies
Dim nomframe As String
Dim b As Boolean

    On Error Resume Next

    ' *** posa els noms del frames, tant si son de grid com si no ***
    Select Case NumTabMto
        Case 0: nomframe = "FrameAux0" 'tarjetas
        Case 1: nomframe = "FrameAux1" 'matriculas
        Case 2: nomframe = "FrameAux2" 'margenes
    End Select
    
    If DatosOkLlin(nomframe) Then
        TerminaBloquear
        If InsertarDesdeForm2(Me, 2, nomframe) Then
            b = BLOQUEADesdeFormulario2(Me, Data1, 1)
            Select Case NumTabMto
                Case 0, 1, 2 ' *** els index de les llinies en grid (en o sense tab) ***
                     CargaGrid NumTabMto, True
                    If b Then BotonAnyadirLinea NumTabMto
            End Select
           
            SituarTab (NumTabMto + 1)
        End If
    End If
End Sub

Private Function ModificarLinea() As Boolean
'Modifica registre en les taules de Llínies
Dim nomframe As String
Dim V As Integer
    
    On Error Resume Next

    ' *** posa els noms del frames, tant si son de grid com si no ***
    Select Case NumTabMto
        Case 0: nomframe = "FrameAux0" 'tarjetas
        Case 1: nomframe = "FrameAux1" 'Matriculas
        Case 2: nomframe = "FrameAux2" 'Margenes
    End Select
    ModificarLinea = False
    If DatosOkLlin(nomframe) Then
        TerminaBloquear
        If ModificaDesdeFormulario2(Me, 2, nomframe) Then
            ModoLineas = 0
            
            V = AdoAux(NumTabMto).Recordset.Fields(1) 'el 2 es el nº de llinia
            CargaGrid NumTabMto, True
            
            ' *** si n'hi han tabs ***
            SituarTab (NumTabMto + 1)

            ' *** si n'hi han tabs que no tenen datagrid, posar el if ***
            PonerFocoGrid Me.DataGridAux(NumTabMto)
            AdoAux(NumTabMto).Recordset.Find (AdoAux(NumTabMto).Recordset.Fields(1).Name & " =" & V)
            
            LLamaLineas NumTabMto, 0
            ModificarLinea = True
        End If
    End If
End Function

Private Function ObtenerWhereCab(conW As Boolean) As String
Dim vWhere As String
    
    vWhere = ""
    If conW Then vWhere = " WHERE "
    ' *** canviar-ho per la clau primaria de la capçalera ***
    vWhere = vWhere & " codsocio=" & Val(text1(0).Text)
    
    ObtenerWhereCab = vWhere
End Function

'' *** neteja els camps dels tabs de grid que
''estan fora d'este, i els camps de descripció ***
Private Sub LimpiarCamposFrame(Index As Integer)
    On Error Resume Next
 
'    Select Case Index
'        Case 0 'Cuentas Bancarias
'            txtAux(11).Text = ""
'            txtAux(12).Text = ""
'        Case 1 'Departamentos
'            txtAux(21).Text = ""
'            txtAux(22).Text = ""
'            txtAux2(22).Text = ""
'            txtAux(23).Text = ""
'            txtAux(24).Text = ""
'        Case 2 'Tarjetas
'            txtAux(50).Text = ""
'            txtAux(51).Text = ""
'        Case 4 'comisiones
'            txtAux2(2).Text = ""
'    End Select
'
    If Err.Number <> 0 Then Err.Clear
End Sub

'' ### [DavidV] 26/04/2006: Activar/desactivar la rueda del ratón.
'Private Sub DataGridAux_GotFocus(Index As Integer)
'  WheelHook DataGridAux(Index)
'End Sub
'Private Sub DataGridAux_LostFocus(Index As Integer)
'  WheelUnHook
'End Sub


'####################################
'
'            C R M
'
'####################################

Private Sub Toolbar3_ButtonClick(ByVal Button As MSComctlLib.Button)
    If Button.Tag = "" Then Exit Sub
    LabelCRM.Caption = ""
    'Levantamos todos los botones y dejamos pulsado el de ahora
    For NumRegElim = 1 To Toolbar3.Buttons.Count
        If Toolbar3.Buttons(NumRegElim).Tag <> "" Then
            If Toolbar3.Buttons(NumRegElim).Index <> Button.Index Then Toolbar3.Buttons(NumRegElim).Value = tbrUnpressed
        End If
    Next NumRegElim
    CargaColumnasCRM CByte(Button.Tag)
    
    'Hacemos las acciones
    If Modo = 2 Then CargaDatosLWCRM
End Sub





Private Sub CargaColumnasCRM(OpcionList As Byte)
Dim Columnas As String
Dim Ancho As String
Dim Alinea As String
Dim Formato As String
Dim Ncol As Integer
Dim c As ColumnHeader
Dim Ordena As Integer
    'Las llamadas cogera las llamadas recibidas desde sllama y las efectuadas desde acciones comerciales con tipoaccion=1
    'para poder ordenarlas tendremos una columna viiblefalse con yyymmddhhmmss
    Ordena = -1
    Select Case OpcionList
    Case 0
        'Acciones comerciales
        LabelCRM.Caption = "Acciones comerciales"
        
        Columnas = "Fecha|Usuario|Estado|Medio|Tipo|Descripcion|"   'fechora ,usuario,estado ,scrmacciones.medio ,tipo,denominacion
        Ancho = "2100|1000|1200|1200|800|2300|"
        'vwColumnRight =1  left=0   center=2
        Alinea = "0|0|0|0|1|0|"
        'Formatos
        Formato = "dd/mm/yyyy hh:mm:ss||||0000||"
        Ncol = 6
               
    Case 1
        'Llamadas
        LabelCRM.Caption = "Llamadas "
        
        Columnas = "Fecha|Usuario|Tipo/Trab|Observaciones|Orden|"   'fechora ,usuario,estado ,scrmacciones.medio ,tipo,denominacion
        Ancho = "2100|1000|1400|4000|0|"
        'vwColumnRight =1  left=0   center=2
        Alinea = "0|0|0|0|0|"
        'Formatos
        Formato = "dd/mm/yyyy hh:mm:ss||||"
        Ncol = 5
    
        Ordena = 5
        
    Case 2
        LabelCRM.Caption = "E-mail"
        Columnas = "Fecha|Enviado|Email|Asunto|Adj|entryID|"  'select fechahora, if(enviado=1,"Enviado","Recibido"),email,asunto,if(adjuntos<>"","*","")  from scrmmail
        Ancho = "1800|825|2565|3899|495|0|"
        'vwColumnRight =1  left=0   center=2
        Alinea = "0|0|0|0|0|"
        'Formatos
        Formato = "dd/mm/yyyy hh:mm||||||"
        Ncol = 6
    
    Case 3
        'COBROS
        LabelCRM.Caption = "Cobros pendientes"
        Columnas = "Fecha Vto.|Factura|Fecha factura|Forma pago|Pendiente|"  'select fechahora, if(enviado=1,"Enviado","Recibido"),email,asunto,if(adjuntos<>"","*","")  from scrmmail
        Ancho = "1600|1500|1300|2400|1495|"
        'vwColumnRight =1  left=0   center=2
        Alinea = "0|1|0|0|1|"
        'Formatos
        Formato = "dd/mm/yyyy||dd/mm/yyyy||" & FormatoImporte & "|"
        Ncol = 5
        
    Case 4
        'COBROS
        LabelCRM.Caption = "Observaciones departamento"
        Columnas = "Departamento|Fecha|Observaciones||"  'select fechahora, if(enviado=1,"Enviado","Recibido"),email,asunto,if(adjuntos<>"","*","")  from scrmmail
        Ancho = "1600|1500|6500|0|"
        'vwColumnRight =1  left=0   center=2
        Alinea = "0|0|0|0|"
        'Formatos
        Formato = "|dd/mm/yyyy|||"
        Ncol = 4
        
        
    Case 5
        'Reclamaciones
        LabelCRM.Caption = "Reclamaciones"
        Columnas = "Fecha|Factura|Observaciones|Importe|codigo|Tipo|"  'select fechahora, if(enviado=1,"Enviado","Recibido"),email,asunto,if(adjuntos<>"","*","")  from scrmmail
        Ancho = "1600|1500|4000|1500|0|1000|"  'La ultima esta oculta
        'vwColumnRight =1  left=0   center=2
        Alinea = "0|0|0|1|0|0|"
        'Formatos
        Formato = "dd/mm/yyyy|||" & FormatoImporte & "|||"
        Ncol = 6
        
    
    Case 6
        'H I S T O R I A L
        LabelCRM.Caption = "Historial"
        Columnas = "Fecha|Usuario|Trabajador|Observaciones||"
        Ancho = "2100|1000|2000|4200|0|"
        'vwColumnRight =1  left=0   center=2
        Alinea = "0|0|0|0|0|"
        'Formatos
        Formato = "dd/mm/yyyy hh:mm:ss|||||"
        Ncol = 5
        
    
    End Select
    
    
    cmdAccCRM(2).visible = OpcionList = 4 'Or OpcionList = 6
    lwCRM.ColumnHeaders.Clear
    
    'Guardo la opcion en el tag
    lwCRM.Tag = OpcionList & "|" & Ncol & "|"
    
    For NumRegElim = 1 To Ncol
         Set c = lwCRM.ColumnHeaders.Add()
         c.Text = RecuperaValor(Columnas, CInt(NumRegElim))
         c.Width = RecuperaValor(Ancho, CInt(NumRegElim))
         c.Alignment = Val(RecuperaValor(Alinea, CInt(NumRegElim)))
         c.Tag = RecuperaValor(Formato, CInt(NumRegElim))
    Next NumRegElim
    
    If Ordena < 0 Then
        lwCRM.Sorted = False
    Else
        lwCRM.Sorted = True
        lwCRM.SortKey = 4
        lwCRM.SortOrder = lvwDescending
    End If
    
End Sub

Private Sub CargaDatosLWCRM()
Dim c As String
Dim bs As Byte
    bs = Screen.MousePointer
    c = Me.lblIndicador.Caption
    lblIndicador.Caption = "Leyendo " & LabelCRM.Caption
    lblIndicador.Refresh
    CargaDatosLWcrm2
    Me.lblIndicador.Caption = c
    Screen.MousePointer = bs
End Sub

Private Sub CargaDatosLWcrm2()
Dim cad As String
Dim RS As ADODB.Recordset
Dim It As ListItem
Dim ElIcono As Integer
Dim GroupBy As String
Dim Kopc As Byte
Dim MeteIT As Boolean
Dim ConexionConta As Boolean  'Si no es conta es ARIGES( conn)

Dim observaciones As String

    On Error GoTo ECargaDatosLW
    
    If Modo <> 2 Then Exit Sub
    
    For NumRegElim = 1 To Toolbar3.Buttons.Count
        If Toolbar3.Buttons(NumRegElim).Value = tbrPressed Then
            ElIcono = Toolbar3.Buttons(NumRegElim).Image
            Exit For
        End If
    Next
    
    
'    'Fecha incio busquedas
'    Text1(46).Text = Format(imgFec(15).Tag, "dd/mm/yyyy")

    'EL where del codclien     lo lleva cada sql
    Kopc = CByte(RecuperaValor(lwCRM.Tag, 1))
    ConexionConta = False
    Select Case Kopc
    Case 0
        'Acciones comerciales
        cad = "select fechora ,usuario,estado ,scrmacciones.medio ,tipo,denominacion from scrmacciones,scrmtipo WHERE scrmacciones.tipo= scrmtipo.codigo "
        cad = cad & " and codclien=" & Data1.Recordset!codClien & " and tipo > 20"  'las 20 primerasprobablemebne no sepongan aqui
        GroupBy = ""
        BuscaChekc = "fechora"
    Case 1
        'Llamadas
        cad = "select feholla,usuario,nomllama1,observac,date_format(feholla,""%Y%m%d%H%i%s"") from sllama,sllama1  where"
        cad = cad & " sllama.codllama1 = sllama1.codllama1"
        cad = cad & " and codclien=" & Data1.Recordset!codClien
        GroupBy = ""
        BuscaChekc = "feholla"
    Case 2
        'eMAIL
        cad = "select fechahora, if(enviado=1,""Enviado"",""Recibido""),email,asunto,"
        cad = cad & "if(adjuntos<>"""",""*"","""") ,entryID from scrmmail"
        cad = cad & " WHERE codclien=" & Data1.Recordset!codClien
        GroupBy = ""
        BuscaChekc = "fechahora"
    Case 3
        'Cobros pendientes
        cad = "SELECT fecvenci,concat(numserie,right(concat(""00000000"",codfaccl),7)),fecfaccl,nomforpa,"
        cad = cad & "impvenci+if(gastos is null,0,gastos)-if(impcobro is null,0,impcobro)  tot"
        cad = cad & " FROM  scobro INNER JOIN sforpa ON scobro.codforpa=sforpa.codforpa "
        cad = cad & " WHERE scobro.codmacta = '" & text1(9).Text & "' "
        
        BuscaChekc = "fecvenci"
        ConexionConta = True
    Case 4
        'Observaciones departamento
        cad = "select if(dpto=1,""Administracion"",if(dpto=2,""Comercial"",if(dpto=3,""SAT"",""Dirección""))),fecha,observa,dpto from scrmobsclien"
        cad = cad & " WHERE codclien=" & Data1.Recordset!codClien
        BuscaChekc = "dpto"
    Case 5
        'Reclamaciones
        'Cobros pendientes
        cad = "select fecreclama,concat(numserie,right(concat(""00000000"",codfaccl),7)),observaciones,impvenci,codigo, case carta When 0 then 'Carta' When 1 then 'EMail' when 2 then 'Telefono' end "
        cad = cad & " from shcocob where codmacta='" & text1(9).Text & "' "
        BuscaChekc = "fecreclama desc ,codigo "
        ConexionConta = True
    Case 6
        'Historial
        cad = "select fechora , usuario, nomtraba, observaciones, date_format(fechora,""%Y%m%d%H%i%s"") fecha from"
        cad = cad & " scrmacciones left join straba on scrmacciones.codtraba=straba.codtraba "
        cad = cad & " WHERE scrmacciones.tipo=2  and codclien= " & Data1.Recordset!codsocio   '2 DE historial
        GroupBy = ""
        BuscaChekc = "fechora"
    End Select
    
    'El group by
    If GroupBy <> "" Then cad = cad & " GROUP BY " & GroupBy
    
    'El ORDER BY
    cad = cad & " ORDER BY " & BuscaChekc
    If Kopc <> 4 Then cad = cad & " DESC"

    
    BuscaChekc = ""
    
    lwCRM.ListItems.Clear
   
    Set RS = New ADODB.Recordset
    If Not ConexionConta Then
        'Conn  ariges
        RS.Open cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Else
        'Va contra la contabilidad  connconta
        RS.Open cad, ConnConta, adOpenKeyset, adLockPessimistic, adCmdText
    End If
    While Not RS.EOF
        If Kopc <> 3 Then
            MeteIT = True
        Else
            If RS!tot <> 0 Then
                MeteIT = True
            Else
                MeteIT = False
            End If
        End If
        On Error Resume Next
        If Kopc = 6 Then observaciones = RS.Fields(3)
        On Error GoTo ECargaDatosLW
        If MeteIT Then
                Set It = lwCRM.ListItems.Add()
                 
                If lwCRM.ColumnHeaders(1).Tag <> "" Then
                    It.Text = Format(RS.Fields(0), lwCRM.ColumnHeaders(1).Tag)
                Else
                    It.Text = RS.Fields(0)
                End If
                'El resto de cmpos
                For NumRegElim = 2 To CInt(RecuperaValor(lwCRM.Tag, 2))
'                    If Kopc = 6 And NumRegElim = 4 Then Stop
'                    If Kopc = 5 And NumRegElim = 4 Then Stop
'
                    If IsNull(RS.Fields(NumRegElim - 1)) Then
                        It.SubItems(NumRegElim - 1) = " "
                    Else
                    
                        If lwCRM.ColumnHeaders(NumRegElim).Tag <> "" Then
                            It.SubItems(NumRegElim - 1) = Format(RS.Fields(NumRegElim - 1), lwCRM.ColumnHeaders(NumRegElim).Tag)
                        Else
                        
                            
                            'Cad = RS.Fields(NumRegElim - 1)
                            cad = DBLetMemo(RS.Fields(NumRegElim - 1))
                            'no TIENE FORMATO. sEGUN LO QUE SEA puedo hacer unas cosas u otras
                            If NumRegElim = 4 And Kopc = 1 Then cad = Replace(cad, vbCrLf, " ")
                            'para las observaciones de la reclamacion tb quito los vbcrlf
                            If NumRegElim = 3 And Kopc = 5 Then cad = Replace(cad, vbCrLf, " ")
                            
                            'Medio
                            If NumRegElim = 3 And Kopc = 0 Then DevuelveMedio cad
                            If NumRegElim = 3 And Kopc = 4 Then cad = Replace(cad, vbCrLf, " ")
                            
                            If Kopc = 6 And NumRegElim = 4 Then cad = observaciones
                            
                            It.SubItems(NumRegElim - 1) = cad
                        
                        End If
                    End If
                Next
                'El icono
                If Kopc = 1 Then
                    It.SmallIcon = 27
                ElseIf Kopc = 2 Then

                    If RS.Fields(1) = "Enviado" Then
                        It.SmallIcon = 28
                    Else
                        It.SmallIcon = 29
                    End If
                Else
                    'el resto ponemos el del toolbar
                    It.SmallIcon = ElIcono
                End If
        End If
        
        
    
        RS.MoveNext
    Wend
    RS.Close
    
    
    If Kopc = 1 Then
        'Llamadas. Las efectuadas las hago desde este punto
        cad = "select fechora ,usuario,nomtraba ,observaciones,date_format(fechora,""%Y%m%d%H%i%s"") from"
        cad = cad & " scrmacciones left join straba on scrmacciones.codtraba=straba.codtraba "
        cad = cad & " WHERE scrmacciones.tipo=1  and codclien= " & Data1.Recordset!codsocio
        RS.Open cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        While Not RS.EOF
            '
            'Coje datos desde dos tablas
            Set It = lwCRM.ListItems.Add()
            It.Text = Format(RS.Fields(0), lwCRM.ColumnHeaders(1).Tag)
           
            For NumRegElim = 2 To CInt(RecuperaValor(lwCRM.Tag, 2))
                If IsNull(RS.Fields(NumRegElim - 1)) Then
                    It.SubItems(NumRegElim - 1) = " "
                Else
                
                    If lwCRM.ColumnHeaders(NumRegElim).Tag <> "" Then
                        It.SubItems(NumRegElim - 1) = Format(RS.Fields(NumRegElim - 1), lwCRM.ColumnHeaders(NumRegElim).Tag)
                    Else
                    
                        
                        cad = RS.Fields(NumRegElim - 1)
                        'no TIENE FORMATO. sEGUN LO QUE SEA puedo hacer unas cosas u otras
                        If NumRegElim = 4 And Kopc = 1 Then cad = Replace(cad, vbCrLf, " ")
  
                        It.SubItems(NumRegElim - 1) = cad
                    
                        
                        
                    End If
                End If
            Next
            It.SmallIcon = 10 '26
            RS.MoveNext
        Wend
        RS.Close
    End If
    Set RS = Nothing
    
    
    Exit Sub
ECargaDatosLW:
    MuestraError Err.Number
    Set RS = Nothing
    
End Sub

Private Sub DevuelveMedio(ByRef cad As String)
    'pendiente,en curso finalizada
    If cad = "0" Then
        cad = "Pendiente"
    ElseIf cad = "1" Then
        cad = "En curso"
    Else
        cad = "Finalizada"
    End If
End Sub

Private Sub ImagenesNavegacion()
    
    SSTab1.TabVisible(4) = vParamAplic.HayCRM
    If vParamAplic.HayCRM Then
    
        With Me.Toolbar3
            .HotImageList = frmPpal.imgListComun_OM
            .DisabledImageList = frmPpal.imgListComun_BN
            .ImageList = frmPpal.imgListComun
'            .ImageList = frmPpal.imgListPpal
            .Buttons(1).Image = 3 '3
            .Buttons(3).Image = 3 '30
            .Buttons(5).Image = 3 '25
            .Buttons(7).Image = 13 '13
            .Buttons(9).Image = 13 '31
            .Buttons(11).Image = 14 '32
            .Buttons(13).Image = 15 '33
        End With
        
        Set lwCRM.SmallIcons = frmPpal.imgListComun 'frmPpal.imgListPpal
        
    End If

End Sub



Private Sub ModificarDatosCuentaContable()
Dim SQL As String
Dim cad As String

    On Error GoTo eModificarDatosCuentaContable


    If text1(1).Text <> NombreAnt Or text1(18).Text <> BancoAnt Or text1(19).Text <> SucurAnt Or text1(20).Text <> DigitoAnt Or text1(21).Text <> CuentaAnt Or _
       DirecAnt <> text1(3).Text Or cPostalAnt <> text1(4).Text Or PoblaAnt <> text1(5).Text Or ProviAnt <> text1(6).Text Or NifAnt <> text1(2).Text Or _
       EMaiAnt <> text1(14).Text Or WebAnt <> text1(10).Text Or _
       IbanAnt <> text1(32).Text Then
        
        cad = "Se han producido cambios en datos del Cliente. " '& vbCrLf
        
'        If NombreAnt <> Text1(2).Text Then Cad = Cad & " Nombre,"
'        If DirecAnt <> Text1(4).Text Then Cad = Cad & " Direccion,"
'        If cPostalAnt <> Text1(5).Text Then Cad = Cad & " CPostal,"
'        If PoblaAnt <> Text1(18).Text Then Cad = Cad & " Población,"
'        If ProviAnt <> Text1(22).Text Then Cad = Cad & " Provincia,"
'        If NifAnt <> Text1(3).Text Then Cad = Cad & " NIF,"
''        If EMaiAnt <> Text1(12).Text Then Cad = Cad & " EMail,"
'        If BancoAnt <> Text1(1).Text Then Cad = Cad & " Banco,"
'        If SucurAnt <> Text1(28).Text Then Cad = Cad & " Sucursal,"
'        If DigitoAnt <> Text1(29).Text Then Cad = Cad & " Dig.Control,"
'        If CuentaAnt <> Text1(30).Text Then Cad = Cad & " Cuenta banco,"
'
'        Cad = Mid(Cad, 1, Len(Cad) - 1)
        
        cad = cad & vbCrLf & vbCrLf & "¿ Desea actualizarlos en la Contabilidad ?" & vbCrLf & vbCrLf
        
        If MsgBox(cad, vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
                        
            SQL = "update cuentas set nommacta = " & DBSet(Trim(text1(1).Text), "T")
            SQL = SQL & ", razosoci = " & DBSet(Trim(text1(1).Text), "T")
            SQL = SQL & ", dirdatos = " & DBSet(Trim(text1(3).Text), "T")
            SQL = SQL & ", codposta = " & DBSet(Trim(text1(4).Text), "T")
            SQL = SQL & ", despobla = " & DBSet(Trim(text1(5).Text), "T")
            SQL = SQL & ", desprovi = " & DBSet(Trim(text1(6).Text), "T")
            SQL = SQL & ", nifdatos = " & DBSet(Trim(text1(2).Text), "T")
            SQL = SQL & ", maidatos = " & DBSet(Trim(text1(14).Text), "T")
            SQL = SQL & ", webdatos = " & DBSet(Trim(text1(10).Text), "T")
            SQL = SQL & ", entidad = " & DBSet(Trim(text1(18).Text), "T", "S")
            SQL = SQL & ", oficina = " & DBSet(Trim(text1(19).Text), "T", "S")
            SQL = SQL & ", cc = " & DBSet(Trim(text1(20).Text), "T", "S")
            SQL = SQL & ", cuentaba = " & DBSet(Trim(text1(21).Text), "T", "S")
            
            '[Monica]22/11/2013: tema iban
            If vEmpresa.HayNorma19_34Nueva = 1 Then
                SQL = SQL & ", iban = " & DBSet(Trim(text1(32).Text), "T", "S")
            End If
            
            SQL = SQL & " where codmacta = " & DBSet(Trim(text1(9).Text), "T")
                        
            ConnConta.Execute SQL
                        
'            MsgBox "Datos de Cuenta modificados correctamente.", vbExclamation
                        
        End If
    End If
    
'   AQUI NO, PORQUE SE PUEDE ESTAR FACTURANDO POR TARJETA
'
'    '[Monica]17/01/2014: modificamos los datos de tesoreria sobre los cobros y pagos pendientes
'    If Text1(18).Text <> BancoAnt Or Text1(19).Text <> SucurAnt Or Text1(20).Text <> DigitoAnt Or Text1(21).Text <> CuentaAnt _
'        Or Text1(32).Text <> IbanAnt Then
'        cad = "Se han producido cambios en la Cta.Bancaria del cliente."
'        cad = cad & vbCrLf & vbCrLf & "¿ Desea actualizar los Cobros y Pagos pendientes en Tesoreria ?" & vbCrLf & vbCrLf
'
'        If HayCobrosPagosPendientes(Text1(9).Text) Then
'            If MsgBox(cad, vbQuestion + vbYesNo + vbDefaultButton1) = vbYes Then
'                If ActualizarCobrosPagosPdtes(Text1(9), Text1(18).Text, Text1(19).Text, Text1(20).Text, Text1(21).Text, Text1(32).Text) Then
''                    MsgBox "Datos en Tesoreria modificados correctamente.", vbExclamation
'                End If
'            End If
'        End If
'    End If
    
    Exit Sub
    
eModificarDatosCuentaContable:
    MuestraError Err.Number, "Modificar Datos Cuenta Contable", Err.Description
End Sub

Private Function HayCobrosPagosPendientes(vCodmacta As String) As Boolean
Dim SQL As String
Dim Sql2 As String
Dim RS As ADODB.Recordset
Dim nRegs As Long

    On Error GoTo eHayCobrosPagosPendientes


    SQL = "select count(*) from scobro where codmacta = " & DBSet(vCodmacta, "T")
    SQL = SQL & " and (codrem is null or codrem = 0) and (transfer is null or transfer = 0) "
    
    Set RS = New ADODB.Recordset
    RS.Open SQL, ConnConta, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not RS.EOF Then
        If DBLet(RS.Fields(0).Value) <> 0 Then nRegs = DBLet(RS.Fields(0).Value)
    End If
            
    SQL = "select count(*) from spagop where ctaprove = " & DBSet(vCodmacta, "T")
    SQL = SQL & " and (transfer is null or transfer = 0)"
    
    Set RS = Nothing
    
    Set RS = New ADODB.Recordset
    RS.Open SQL, ConnConta, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not RS.EOF Then
        If DBLet(RS.Fields(0).Value) <> 0 Then nRegs = nRegs + DBLet(RS.Fields(0).Value)
    End If
    Set RS = Nothing
    
    HayCobrosPagosPendientes = (nRegs <> 0)
    Exit Function
    
eHayCobrosPagosPendientes:
    MuestraError Err.Number, "Hay Cobros Pagos Pendientes", Err.Description
End Function

Private Function ActualizarCobrosPagosPdtes(vCodmacta As String, vBanco As String, vSucur As String, vDigcon As String, vCta As String, vIban As String) As Boolean
Dim Sql2 As String
    
    On Error GoTo eActualizarCobrosPagosPdtes
    
    ConnConta.BeginTrans
    
    ActualizarCobrosPagosPdtes = False
    
    Sql2 = "update scobro set codbanco = " & DBSet(vBanco, "N", "S") & ", codsucur = " & DBSet(vSucur, "N", "S")
    Sql2 = Sql2 & ", digcontr = " & DBSet(vDigcon, "T", "S") & ", cuentaba = " & DBSet(vCta, "T", "S")
    
    '[Monica]22/11/2013: tema iban
    If vEmpresa.HayNorma19_34Nueva = 1 Then
        Sql2 = Sql2 & ", iban = " & DBSet(vIban, "T", "S")
    End If
    
    Sql2 = Sql2 & " where codmacta = " & DBSet(vCodmacta, "T")
    Sql2 = Sql2 & " and (codrem is null or codrem = 0) and (transfer is null or transfer = 0)"
    
    ConnConta.Execute Sql2
    
    Sql2 = "update spagop set entidad = " & DBSet(vBanco, "T", "S") & ", oficina = " & DBSet(vSucur, "T", "S")
    Sql2 = Sql2 & ", cc = " & DBSet(vDigcon, "T", "S") & ", cuentaba = " & DBSet(vCta, "T", "S")
    
    '[Monica]22/11/2013: tema iban
    If vEmpresa.HayNorma19_34Nueva = 1 Then
        Sql2 = Sql2 & ", iban = " & DBSet(vIban, "T", "S")
    End If
    
    Sql2 = Sql2 & " where ctaprove = " & DBSet(vCodmacta, "T")
    Sql2 = Sql2 & " and (transfer is null or transfer = 0)"
   
    ConnConta.Execute Sql2
    
    ActualizarCobrosPagosPdtes = True
    ConnConta.CommitTrans
    Exit Function
    
eActualizarCobrosPagosPdtes:
    ConnConta.RollbackTrans
    MuestraError Err.Number, "Actualizar Cobros Pagos Pendientes", Err.Description
End Function


