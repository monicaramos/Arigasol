VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
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
      TabIndex        =   54
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
         TabIndex        =   56
         Top             =   240
         Width           =   660
      End
      Begin VB.Label Label1 
         Caption         =   "Código Cliente"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   55
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Height          =   555
      Index           =   0
      Left            =   240
      TabIndex        =   51
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
         TabIndex        =   52
         Top             =   180
         Width           =   2655
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   10410
      TabIndex        =   48
      Top             =   6960
      Width           =   1035
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   9120
      TabIndex        =   47
      Top             =   6960
      Width           =   1035
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5460
      Left            =   240
      TabIndex        =   53
      Top             =   1320
      Width           =   11295
      _ExtentX        =   19923
      _ExtentY        =   9631
      _Version        =   393216
      Style           =   1
      Tabs            =   4
      TabsPerRow      =   7
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
      Tab(0).Control(13)=   "text1(2)"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "text1(3)"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "text1(6)"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "text1(5)"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "text2(7)"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "text1(7)"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "text2(8)"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "text1(8)"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "text1(4)"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "FrameDatosAlta"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "FrameDatosContacto"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "text1(24)"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "text1(9)"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "text2(9)"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "text1(25)"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).ControlCount=   28
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
      Begin VB.Frame FrameAux2 
         BorderStyle     =   0  'None
         Height          =   4215
         Left            =   -74760
         TabIndex        =   96
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
            TabIndex        =   100
            Tag             =   "Euros/litro|N|S|||smargen|euroslitro|#,##0.000||"
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
            TabIndex        =   105
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
            TabIndex        =   104
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
            TabIndex        =   101
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
            TabIndex        =   99
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
            TabIndex        =   98
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
            TabIndex        =   97
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
            TabIndex        =   102
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
            Bindings        =   "frmManClien.frx":007C
            Height          =   3645
            Index           =   2
            Left            =   0
            TabIndex        =   103
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
         TabIndex        =   90
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
            TabIndex        =   33
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
            TabIndex        =   35
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
            TabIndex        =   36
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
            TabIndex        =   34
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
            TabIndex        =   91
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
            Bindings        =   "frmManClien.frx":0094
            Height          =   3645
            Index           =   1
            Left            =   0
            TabIndex        =   92
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
         TabIndex        =   87
         Top             =   480
         Width           =   11055
         Begin VB.TextBox txtAux 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   290
            Index           =   12
            Left            =   6840
            MaxLength       =   10
            TabIndex        =   46
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
            TabIndex        =   41
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
            TabIndex        =   37
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
            TabIndex        =   38
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
            Left            =   4560
            MaxLength       =   4
            TabIndex        =   42
            Tag             =   "Banco|N|S|0|9999|starje|codbanco|0000||"
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
            Left            =   5280
            MaxLength       =   4
            TabIndex        =   43
            Tag             =   "Oficina|N|S|0|9999|starje|codsucur|0000||"
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
            TabIndex        =   44
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
            TabIndex        =   45
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
            TabIndex        =   39
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
            ItemData        =   "frmManClien.frx":00AC
            Left            =   1200
            List            =   "frmManClien.frx":00AE
            Style           =   2  'Dropdown List
            TabIndex        =   40
            Tag             =   "Tipo Tarjeta|N|N|||starje|tiptarje|||"
            Top             =   3720
            Visible         =   0   'False
            Width           =   615
         End
         Begin MSComctlLib.Toolbar ToolAux 
            Height          =   390
            Index           =   0
            Left            =   0
            TabIndex        =   88
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
            Bindings        =   "frmManClien.frx":00B0
            Height          =   3825
            Index           =   0
            Left            =   0
            TabIndex        =   89
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
      End
      Begin VB.TextBox text2 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   9
         Left            =   2700
         TabIndex        =   78
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
         Height          =   1395
         Index           =   24
         Left            =   5880
         MaxLength       =   250
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   32
         Tag             =   "Observaciones|T|S|||ssocio|obssocio|||"
         Top             =   3870
         Width           =   5295
      End
      Begin VB.Frame FrameDatosContacto 
         Caption         =   "Datos Contacto"
         ForeColor       =   &H00972E0B&
         Height          =   1800
         Left            =   225
         TabIndex        =   66
         Top             =   3510
         Width           =   5415
         Begin VB.TextBox text1 
            Height          =   285
            Index           =   10
            Left            =   1200
            MaxLength       =   40
            TabIndex        =   11
            Tag             =   "Web|T|S|||ssocio|wwwsocio|||"
            Top             =   360
            Width           =   4095
         End
         Begin VB.TextBox text1 
            Height          =   285
            Index           =   11
            Left            =   1200
            MaxLength       =   10
            TabIndex        =   12
            Tag             =   "Teléfono|T|S|||ssocio|telsocio|||"
            Top             =   730
            Width           =   1455
         End
         Begin VB.TextBox text1 
            Height          =   285
            Index           =   12
            Left            =   3660
            MaxLength       =   10
            TabIndex        =   13
            Tag             =   "Móvil|T|S|||ssocio|movsocio|||"
            Top             =   730
            Width           =   1455
         End
         Begin VB.TextBox text1 
            Height          =   285
            Index           =   13
            Left            =   1200
            MaxLength       =   10
            TabIndex        =   14
            Tag             =   "Fax|T|S|||ssocio|faxsocio|||"
            Top             =   1100
            Width           =   1455
         End
         Begin VB.TextBox text1 
            Height          =   285
            Index           =   14
            Left            =   1200
            MaxLength       =   40
            TabIndex        =   15
            Tag             =   "E-mail|T|S|||ssocio|maisocio|||"
            Top             =   1440
            Width           =   4095
         End
         Begin VB.Image imgWeb 
            Height          =   240
            Index           =   0
            Left            =   735
            Picture         =   "frmManClien.frx":00C8
            Top             =   360
            Width           =   240
         End
         Begin VB.Label Label9 
            Caption         =   "Web"
            Height          =   255
            Left            =   240
            TabIndex        =   74
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
            TabIndex        =   70
            Top             =   730
            Width           =   735
         End
         Begin VB.Label Label12 
            Caption         =   "Móvil"
            Height          =   255
            Left            =   3165
            TabIndex        =   69
            Top             =   730
            Width           =   495
         End
         Begin VB.Label Label14 
            Caption         =   "Fax"
            Height          =   255
            Left            =   240
            TabIndex        =   68
            Top             =   1100
            Width           =   735
         End
         Begin VB.Label Label16 
            Caption         =   "E-mail"
            Height          =   255
            Left            =   240
            TabIndex        =   67
            Top             =   1470
            Width           =   495
         End
      End
      Begin VB.Frame FrameDatosAlta 
         Caption         =   "Datos Facturación"
         ForeColor       =   &H00972E0B&
         Height          =   3195
         Left            =   5760
         TabIndex        =   63
         Top             =   360
         Width           =   5415
         Begin VB.CheckBox chkAux 
            Caption         =   "Envio Factura por eMail"
            Height          =   195
            Index           =   0
            Left            =   150
            TabIndex        =   31
            Tag             =   "Envio Factura eMail|N|N|0|1|ssocio|envfactemail|||"
            Top             =   2820
            Width           =   2295
         End
         Begin VB.ComboBox Combo1 
            Height          =   315
            Index           =   5
            ItemData        =   "frmManClien.frx":0652
            Left            =   1290
            List            =   "frmManClien.frx":065C
            Style           =   2  'Dropdown List
            TabIndex        =   28
            Tag             =   "Factura con FP|N|N|0|1|ssocio|facturafp|||"
            Top             =   2385
            Width           =   1095
         End
         Begin VB.ComboBox Combo1 
            Height          =   315
            Index           =   4
            ItemData        =   "frmManClien.frx":0672
            Left            =   3720
            List            =   "frmManClien.frx":067C
            Style           =   2  'Dropdown List
            TabIndex        =   29
            Tag             =   "Grupo Est.Artículos|N|N|0|1|ssocio|grupoestartic|||"
            Top             =   2370
            Width           =   1305
         End
         Begin VB.TextBox text1 
            Height          =   285
            Index           =   21
            Left            =   3330
            MaxLength       =   10
            TabIndex        =   22
            Tag             =   "Cuenta|T|S|||ssocio|cuentaba|||"
            Top             =   948
            Width           =   1680
         End
         Begin VB.TextBox text1 
            Height          =   285
            Index           =   20
            Left            =   2730
            MaxLength       =   2
            TabIndex        =   21
            Tag             =   "D.C.|T|S|||ssocio|digcontr|||"
            Top             =   948
            Width           =   480
         End
         Begin VB.TextBox text1 
            Height          =   285
            Index           =   19
            Left            =   2010
            MaxLength       =   4
            TabIndex        =   20
            Tag             =   "Sucursal|T|S|||ssocio|codsucur|||"
            Top             =   948
            Width           =   600
         End
         Begin VB.TextBox text1 
            Height          =   285
            Index           =   18
            Left            =   1290
            MaxLength       =   4
            TabIndex        =   19
            Tag             =   "Banco|T|S|||ssocio|codbanco|||"
            Top             =   948
            Width           =   600
         End
         Begin VB.TextBox text1 
            Height          =   285
            Index           =   23
            Left            =   3720
            MaxLength       =   8
            TabIndex        =   24
            Tag             =   "Dto./Litro|N|N|0|9.9999|ssocio|dtolitro|0.0000||"
            Top             =   1320
            Width           =   1290
         End
         Begin VB.TextBox text1 
            Height          =   285
            Index           =   22
            Left            =   1290
            MaxLength       =   1
            TabIndex        =   23
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
            TabIndex        =   82
            Top             =   600
            Width           =   3135
         End
         Begin VB.TextBox text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   17
            Left            =   1290
            MaxLength       =   4
            TabIndex        =   18
            Tag             =   "Código F.Pago|N|N|0|999|ssocio|codforpa|000||"
            Top             =   594
            Width           =   555
         End
         Begin VB.ComboBox Combo1 
            Height          =   315
            Index           =   3
            ItemData        =   "frmManClien.frx":0692
            Left            =   3720
            List            =   "frmManClien.frx":069C
            Style           =   2  'Dropdown List
            TabIndex        =   27
            Tag             =   "Bonif.Especial|N|N|0|1|ssocio|bonifesp|||"
            Top             =   2010
            Width           =   1305
         End
         Begin VB.ComboBox Combo1 
            Height          =   315
            Index           =   2
            ItemData        =   "frmManClien.frx":06B2
            Left            =   1290
            List            =   "frmManClien.frx":06BC
            Style           =   2  'Dropdown List
            TabIndex        =   26
            Tag             =   "Bonif.Basica|N|N|0|1|ssocio|bonifbas|||"
            Top             =   2010
            Width           =   1095
         End
         Begin VB.ComboBox Combo1 
            Height          =   315
            Index           =   1
            ItemData        =   "frmManClien.frx":06D2
            Left            =   3720
            List            =   "frmManClien.frx":06DC
            Style           =   2  'Dropdown List
            TabIndex        =   30
            Tag             =   "Imprime Factura|N|N|0|1|ssocio|impfactu|||"
            Top             =   1650
            Width           =   1305
         End
         Begin VB.ComboBox Combo1 
            Height          =   315
            Index           =   0
            ItemData        =   "frmManClien.frx":06F2
            Left            =   1290
            List            =   "frmManClien.frx":06FC
            Style           =   2  'Dropdown List
            TabIndex        =   25
            Tag             =   "Tipo de cliente|N|N|0|2|ssocio|tipsocio|||"
            Top             =   1650
            Width           =   1095
         End
         Begin VB.TextBox text1 
            Height          =   285
            Index           =   16
            Left            =   3810
            MaxLength       =   10
            TabIndex        =   17
            Tag             =   "F.Baja|F|S|||ssocio|fechabaj|dd/mm/yyyy||"
            Top             =   240
            Width           =   1200
         End
         Begin VB.TextBox text1 
            Height          =   285
            Index           =   15
            Left            =   1290
            MaxLength       =   10
            TabIndex        =   16
            Tag             =   "F.Alta|F|N|||ssocio|fechaalt|dd/mm/yyyy||"
            Top             =   240
            Width           =   1200
         End
         Begin VB.Label Label2 
            Caption         =   "Fact.F.P ficha"
            Height          =   255
            Index           =   2
            Left            =   135
            TabIndex        =   94
            Top             =   2430
            Width           =   1065
         End
         Begin VB.Label Label2 
            Caption         =   "Grupo Est.Art."
            Height          =   255
            Index           =   1
            Left            =   2550
            TabIndex        =   93
            Top             =   2430
            Width           =   1005
         End
         Begin VB.Image imgFec 
            Height          =   240
            Index           =   16
            Left            =   3510
            Picture         =   "frmManClien.frx":0712
            ToolTipText     =   "Buscar fecha"
            Top             =   240
            Width           =   240
         End
         Begin VB.Image imgFec 
            Height          =   240
            Index           =   15
            Left            =   990
            Picture         =   "frmManClien.frx":079D
            ToolTipText     =   "Buscar fecha"
            Top             =   240
            Width           =   240
         End
         Begin VB.Label Label17 
            Caption         =   "Datos Banco"
            Height          =   255
            Left            =   120
            TabIndex        =   86
            Top             =   975
            Width           =   975
         End
         Begin VB.Label Label15 
            Caption         =   "Dto. x Litro"
            Height          =   255
            Left            =   2550
            TabIndex        =   85
            Top             =   1335
            Width           =   855
         End
         Begin VB.Label Label13 
            Caption         =   "Tarifa"
            Height          =   255
            Left            =   120
            TabIndex        =   84
            Top             =   1335
            Width           =   855
         End
         Begin VB.Label Label11 
            Caption         =   "F.Pago"
            Height          =   255
            Left            =   120
            TabIndex        =   83
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
            Caption         =   "Bonif.Espec."
            Height          =   255
            Index           =   5
            Left            =   2550
            TabIndex        =   81
            Top             =   2040
            Width           =   975
         End
         Begin VB.Label Label2 
            Caption         =   "Bonif.Basica"
            Height          =   255
            Index           =   4
            Left            =   120
            TabIndex        =   80
            Top             =   2040
            Width           =   975
         End
         Begin VB.Label Label2 
            Caption         =   "Imp. Factura"
            Height          =   255
            Index           =   3
            Left            =   2550
            TabIndex        =   77
            Top             =   1680
            Width           =   1155
         End
         Begin VB.Label Label2 
            Caption         =   "Tipo Cliente"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   76
            Top             =   1680
            Width           =   975
         End
         Begin VB.Label Label3 
            Caption         =   "Fecha Baja"
            Height          =   255
            Left            =   2580
            TabIndex        =   75
            Top             =   255
            Width           =   825
         End
         Begin VB.Label Label21 
            Caption         =   "Fecha Alta"
            Height          =   255
            Left            =   120
            TabIndex        =   65
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
         TabIndex        =   50
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
         TabIndex        =   49
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
      Begin VB.Label Label1 
         Caption         =   "Nro.Socio"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   95
         Top             =   3015
         Width           =   870
      End
      Begin VB.Label Label8 
         Caption         =   "Cta.Conta."
         Height          =   255
         Left            =   240
         TabIndex        =   79
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
         Top             =   3585
         Width           =   240
      End
      Begin VB.Label Label29 
         Caption         =   "Observaciones"
         Height          =   255
         Left            =   5880
         TabIndex        =   73
         Top             =   3600
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
         TabIndex        =   62
         Top             =   2280
         Width           =   855
      End
      Begin VB.Label Label18 
         Caption         =   "Colectivo"
         Height          =   255
         Left            =   240
         TabIndex        =   61
         Top             =   1920
         Width           =   735
      End
      Begin VB.Label Label7 
         Caption         =   "Provincia"
         Height          =   255
         Left            =   240
         TabIndex        =   60
         Top             =   1605
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Población"
         Height          =   255
         Index           =   26
         Left            =   240
         TabIndex        =   59
         Top             =   1240
         Width           =   735
      End
      Begin VB.Label Label6 
         Caption         =   "Domicilio"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   58
         Top             =   880
         Width           =   735
      End
      Begin VB.Label Label5 
         Caption         =   "NIF"
         Height          =   255
         Left            =   240
         TabIndex        =   57
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
      TabIndex        =   71
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
         NumButtons      =   21
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
            Object.ToolTipText     =   "Salir"
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Primero"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Anterior"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Siguiente"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button21 {66833FEA-8583-11D1-B16A-00C0F0283628} 
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
         TabIndex        =   72
         Top             =   120
         Width           =   1215
      End
   End
   Begin VB.CommandButton cmdRegresar 
      Caption         =   "&Regresar"
      Height          =   375
      Left            =   10440
      TabIndex        =   64
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
Private WithEvents frmFpa As frmManFpago 'F.Pago
Attribute frmFpa.VB_VarHelpID = -1
Private WithEvents frmCoo As frmManCoope 'Colectivos
Attribute frmCoo.VB_VarHelpID = -1
Private WithEvents frmSit As frmManSitua 'Situaciones
Attribute frmSit.VB_VarHelpID = -1

Private WithEvents frmCtas As frmCtasConta 'cuentas contables
Attribute frmCtas.VB_VarHelpID = -1
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

Private Sub cmbAux_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
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
                    Text2(9).Text = PonerNombreCuenta(Text1(9), Modo, Text1(0).Text)
        
                    Data1.RecordSource = "Select * from " & NombreTabla & Ordenacion
                    PosicionarData
                End If
            Else
                ModoLineas = 0
            End If
            
        Case 4  'MODIFICAR
            If DatosOk Then
                If ModificaDesdeFormulario2(Me, 1) Then
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
Dim i As Integer

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
        .Buttons(15).Image = 11  'Eixir
        'el 13 i el 14 son separadors
        .Buttons(btnPrimero).Image = 6  'Primer
        .Buttons(btnPrimero + 1).Image = 7 'Anterior
        .Buttons(btnPrimero + 2).Image = 8 'Següent
        .Buttons(btnPrimero + 3).Image = 9 'Últim
    End With
    
    ' ******* si n'hi han llínies *******
    'ICONETS DE LES BARRES ALS TABS DE LLÍNIA
    For i = 0 To ToolAux.Count - 1
        With Me.ToolAux(i)
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
    Next i
    ' ***********************************
    
    'cargar IMAGES de busqueda
    For i = 0 To Me.imgBuscar.Count - 1
        Me.imgBuscar(i).Picture = frmPpal.imgListImages16.ListImages(1).Picture
    Next i
    
    'carga IMAGES de mail
    For i = 0 To Me.ImgMail.Count - 1
        Me.ImgMail(i).Picture = frmPpal.imgListImages16.ListImages(2).Picture
    Next i
    
    'IMAGES para zoom
    For i = 0 To Me.imgZoom.Count - 1
        Me.imgZoom(i).Picture = frmPpal.imgListImages16.ListImages(3).Picture
    Next i
    
    
    
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
        Text1(0).BackColor = vbYellow 'codclien
        ' ****************************************************************************
    End If
End Sub

Private Sub LimpiarCampos()
    On Error Resume Next
    
    limpiar Me   'Mètode general: Neteja els controls TextBox
    lblIndicador.Caption = ""
    
    ' *** si n'hi han combos a la capçalera ***
    Me.Combo1(0).ListIndex = -1
    Me.Combo1(1).ListIndex = -1
    Me.Combo1(2).ListIndex = -1
    Me.Combo1(3).ListIndex = -1
    Me.Combo1(4).ListIndex = -1
    ' *****************************************


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
Dim i As Integer, NumReg As Byte
Dim b As Boolean

    On Error GoTo EPonerModo
 
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
    NumReg = 1
    If Not Data1.Recordset.EOF Then
        If Data1.Recordset.RecordCount > 1 Then NumReg = 2 'Només es per a saber que n'hi ha + d'1 registre
    End If
    DesplazamientoVisible Me.Toolbar1, btnPrimero, b, NumReg
    
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
        For i = 0 To 2
            CargaGrid i, False
        Next i
    End If
    
    b = (Modo = 4) Or (Modo = 2)
    For i = 0 To 2
        DataGridAux(i).Enabled = b
    Next i
      
    ' ****** si n'hi han combos a la capçalera ***********************
    If (Modo = 0) Or (Modo = 2) Or (Modo = 5) Then
        Combo1(0).Enabled = False
        Combo1(0).BackColor = &H80000018 'groc
        Combo1(1).Enabled = False
        Combo1(1).BackColor = &H80000018 'groc
        Combo1(2).Enabled = False
        Combo1(2).BackColor = &H80000018 'groc
        Combo1(3).Enabled = False
        Combo1(3).BackColor = &H80000018 'groc
        Combo1(4).Enabled = False
        Combo1(4).BackColor = &H80000018 'groc
    ElseIf (Modo = 1) Or (Modo = 3) Or (Modo = 4) Then
        Combo1(0).Enabled = True
        Combo1(0).BackColor = &H80000005 'blanc
        Combo1(1).Enabled = True
        Combo1(1).BackColor = &H80000005 'blanc
        Combo1(2).Enabled = True
        Combo1(2).BackColor = &H80000005 'blanc
        Combo1(3).Enabled = True
        Combo1(3).BackColor = &H80000005 'blanc
        Combo1(4).Enabled = True
        Combo1(4).BackColor = &H80000005 'blanc
    End If
    ' ****************************************************************
    
    PonerModoOpcionesMenu (Modo) 'Activar opcions menú según modo
    PonerOpcionesMenu   'Activar opcions de menú según nivell
                        'de permisos de l'usuari

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
Dim i As Byte
    
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
    For i = 0 To ToolAux.Count - 1
        ToolAux(i).Buttons(1).Enabled = b
        If b Then bAux = (b And Me.AdoAux(i).Recordset.RecordCount > 0)
        ToolAux(i).Buttons(2).Enabled = bAux
        ToolAux(i).Buttons(3).Enabled = bAux
    Next i
    
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
Dim sql As String
Dim tabla As String
    
    ' ********* si n'hi han tabs, dona igual si en datagrid o no ***********
    Select Case Index
               
        Case 0 'TARJETAS
            sql = "SELECT codsocio,numlinea,numtarje,tiptarje,CASE tiptarje WHEN 0 THEN ""Normal"" WHEN 1 THEN ""Bonificado"" WHEN 2 THEN ""Profesional"" END, nomtarje,codbanco,codsucur,digcontr,cuentaba, matricul "
            sql = sql & " FROM starje "
            If enlaza Then
                sql = sql & ObtenerWhereCab(True)
            Else
                sql = sql & " WHERE starje.codsocio = -1"
            End If
            sql = sql & " ORDER BY starje.numlinea"
               
        Case 1 'MATRICULAS
            sql = "SELECT codsocio,numlinea,matricul,observac "
            sql = sql & " FROM smatri"
            If enlaza Then
                sql = sql & ObtenerWhereCab(True)
            Else
                sql = sql & " WHERE smatri.codsocio = -1"
            End If
            sql = sql & " ORDER BY smatri.numlinea"
            
        Case 2 'MARGENES DE PRECIOS DE ARTICULOS DE COMBUSTIBLE
            sql = "SELECT codsocio,numlinea,smargen.codartic, nomartic, margen, euroslitro "
            sql = sql & " FROM smargen INNER JOIN sartic ON smargen.codartic = sartic.codartic "
            If enlaza Then
                sql = sql & ObtenerWhereCab(True)
            Else
                sql = sql & " WHERE smargen.codsocio = -1"
            End If
            sql = sql & " ORDER BY smargen.numlinea"
            
            
            
    End Select
    
    MontaSQLCarga = sql
End Function

Private Sub frmArt_DatoSeleccionado(CadenaSeleccion As String)
    txtAux(indice).Text = RecuperaValor(CadenaSeleccion, 1)
    FormateaCampo txtAux(indice)
    Text2(7).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmB_Selecionado(CadenaDevuelta As String)
    Dim Aux As String
    
    If CadenaDevuelta <> "" Then
        HaDevueltoDatos = True
        Screen.MousePointer = vbHourglass
        'Sabem quins camps son els que mos torna
        'Creem una cadena consulta i posem els datos
        CadB = ""
        Aux = ValorDevueltoFormGrid(Text1(0), CadenaDevuelta, 1)
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
    Text1(0).Text = vValor 'codcoope
    
    FormateaCampo Text1(0)
'    text2(7).Text = RecuperaValor(CadenaSeleccion, 2) 'nomcoope
        Modo = 1
        cmdAceptar_Click
End Sub

Private Sub frmZ_Actualizar(vCampo As String)
     Text1(indice).Text = vCampo
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
    If Text1(Index).Text <> "" Then frmC.NovaData = Text1(Index).Text
    ' ********************************************

    frmC.Show vbModal
    Set frmC = Nothing
    ' *** repasar si el camp es txtAux o Text1 ***
    PonerFoco Text1(CByte(imgFec(15).Tag)) '<===
    ' ********************************************
End Sub

Private Sub frmC_Selec(vFecha As Date)
    ' *** repasar si el camp es txtAux o Text1 ***
    Text1(CByte(imgFec(15).Tag)).Text = Format(vFecha, "dd/mm/yyyy") '<===
    ' ********************************************
End Sub
' *****************************************************

'Private Sub btnFec_Click(Index As Integer)
'    imgFec_Click (Index)
'End Sub

Private Sub imgMail_Click(Index As Integer)
    If Index = 0 Then
        If Text1(14).Text <> "" Then
            LanzaMailGnral Text1(14).Text
        End If
    End If
End Sub

Private Sub imgZoom_Click(Index As Integer)
    
    Set frmZ = New frmZoom

    If Index = 0 Then
        indice = 24
        frmZ.pTitulo = "Observaciones del Cliente"
        frmZ.pValor = Text1(indice).Text
        frmZ.pModo = Modo
    
        frmZ.Show vbModal
        Set frmZ = Nothing
            
        PonerFoco Text1(indice)
    End If
End Sub

Private Sub mnBuscar_Click()
    BotonBuscar
    Combo1(0).ListIndex = -1 'quan busque, per defecte no seleccione cap tipo de client
    Combo1(1).ListIndex = -1
    Combo1(2).ListIndex = -1
    Combo1(3).ListIndex = -1
    Combo1(4).ListIndex = -1
End Sub

Private Sub mnBuscarTarjeta_Click()
    Set frmTra = New frmTraerTarje
    frmTra.DatosADevolverBusqueda = "0|1|"
    frmTra.CodigoActual = Text1(0).Text
    frmTra.Show vbModal
    Set frmTra = Nothing
    PonerFoco Text1(0)
End Sub

Private Sub mnClientesLibres_Click()
    BotonClientesLibres
End Sub

Private Sub mnEliminar_Click()
    BotonEliminar
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
        Case 15   'Eixir
            mnSalir_Click
            
        Case btnPrimero To btnPrimero + 3 'Fleches Desplaçament
            Desplazamiento (Button.Index - btnPrimero)
    End Select
End Sub

Private Sub BotonBuscar()
Dim i As Integer
' ***** Si la clau primaria de la capçalera no es Text1(0), canviar-ho en <=== *****
    If Modo <> 1 Then
        LimpiarCampos
        PonerModo 1
        PonerFoco Text1(0) ' <===
        Text1(0).BackColor = vbYellow ' <===
        ' *** si n'hi han combos a la capçalera ***
        For i = 0 To Combo1.Count - 1
            Combo1(i).ListIndex = -1
        Next i
    Else
        HacerBusqueda
        If Data1.Recordset.EOF Then
            Text1(kCampo).Text = ""
            Text1(kCampo).BackColor = vbYellow
            PonerFoco Text1(kCampo)
        End If
    End If
' ******************************************************************************
End Sub

Private Sub HacerBusqueda()

    CadB = ObtenerBusqueda2(Me, 1)
    
    If chkVistaPrevia(0) = 1 Then
        MandaBusquedaPrevia CadB
    ElseIf CadB <> "" Then
        CadenaConsulta = "select * from " & NombreTabla & " WHERE " & CadB & " " & Ordenacion
        PonerCadenaBusqueda
    Else
        ' *** foco al 1r camp visible de la capçalera que siga clau primaria ***
        PonerFoco Text1(0)
        ' **********************************************************************
    End If
End Sub

Private Sub MandaBusquedaPrevia(CadB As String)
    Dim cad As String
        
    'Cridem al form
    ' **************** arreglar-ho per a vore lo que es desije ****************
    ' NOTA: el total d'amples de ParaGrid, ha de sumar 100
    cad = ""
    cad = cad & ParaGrid(Text1(0), 15, "Cód.")
    cad = cad & ParaGrid(Text1(1), 60, "Nombre")
    cad = cad & ParaGrid(Text1(2), 25, "N.I.F.")
    If cad <> "" Then
        Screen.MousePointer = vbHourglass
        Set frmB = New frmBuscaGrid
        frmB.vCampos = cad
        frmB.vtabla = NombreTabla
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
            PonerFoco Text1(kCampo)
        End If
    End If
End Sub

Private Sub cmdRegresar_Click()
Dim cad As String
Dim Aux As String
Dim i As Integer
Dim j As Integer

    If Data1.Recordset.EOF Then
        MsgBox "Ningún registro devuelto.", vbExclamation
        Exit Sub
    End If
    
    cad = ""
    i = 0
    Do
        j = i + 1
        i = InStr(j, DatosADevolverBusqueda, "|")
        If i > 0 Then
            Aux = Mid(DatosADevolverBusqueda, j, i - j)
            j = Val(Aux)
            cad = cad & Text1(j).Text & "|"
        End If
    Loop Until i = 0
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
        PonerCampos
        PonerModo 2
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
    Text1(0).Text = SugerirCodigoSiguienteStr("ssocio", "codsocio")
    FormateaCampo Text1(0)
       
    Text1(15).Text = Format(Now, "dd/mm/yyyy") ' Quan afegixc pose en F.Alta i F.Modificación la data actual
    PosicionarCombo Combo1(0), 0
    PosicionarCombo Combo1(1), 0
    PosicionarCombo Combo1(2), 0
    PosicionarCombo Combo1(3), 0
    PosicionarCombo Combo1(4), 0
        
    PonerFoco Text1(0) '*** 1r camp visible que siga PK ***
    
    ' *** si n'hi han camps de descripció a la capçalera ***
    'PosarDescripcions

    ' *** si n'hi han tabs, em posicione al 1r ***
    Me.SSTab1.Tab = 0
End Sub

Private Sub BotonModificar()

    PonerModo 4

    ' *** bloquejar els camps visibles de la clau primaria de la capçalera ***
    BloquearTxt Text1(0), True
    
    ' *** foco al 1r camp visible que NO siga clau primaria ***
    PonerFoco Text1(1)
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
    cad = cad & vbCrLf & "Código: " & Format(Data1.Recordset.Fields(0), FormatoCampo(Text1(0)))
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
Dim i As Integer
Dim codPobla As String, desPobla As String
Dim CPostal As String, desProvi As String, desPais As String

    If Data1.Recordset.EOF Then Exit Sub
    PonerCamposForma2 Me, Data1, 1 'opcio=1: posa el format o els camps de la capçalera
    
    ' *** si n'hi han llínies en datagrids ***
    'For i = 0 To DataGridAux.Count - 1
    For i = 0 To 2
        CargaGrid i, True
        If Not AdoAux(i).Recordset.EOF Then _
            PonerCamposForma2 Me, AdoAux(i), 2, "FrameAux" & i
    Next i

    ' ************* configurar els camps de les descripcions de la capçalera *************
    Text2(7).Text = PonerNombreDeCod(Text1(7), "scoope", "nomcoope", "codcoope", "N")
    Text2(8).Text = PonerNombreDeCod(Text1(8), "ssitua", "nomsitua")
    Text2(17).Text = PonerNombreDeCod(Text1(17), "sforpa", "nomforpa", "codforpa", "N")
    Text2(9).Text = PonerNombreCuenta(Text1(9), Modo)
    ' ********************************************************************************
    
    '-- Esto permanece para saber donde estamos
    lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
    
    
    Me.SSTab1.TabEnabled(3) = (CInt(Data1.Recordset!bonifesp) = 1)
    
    
    PonerModoOpcionesMenu (Modo)
    PonerOpcionesMenu
End Sub

Private Sub cmdCancelar_Click()
Dim i As Integer
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
                PonerFoco Text1(0)

        Case 4  'Modificar
                TerminaBloquear
                PonerModo 2
                PonerCampos
                ' *** primer camp visible de la capçalera ***
                PonerFoco Text1(0)
        
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
End Sub

Private Function DatosOk() As Boolean
Dim b As Boolean
'Dim Datos As String

    On Error GoTo EDatosOK

    DatosOk = False
    b = CompForm2(Me, 1)
    If Not b Then Exit Function
    
    ' *** canviar els arguments de la funcio, el mensage i repasar si n'hi ha codEmpre ***
    If (Modo = 3) Then 'insertar
        'comprobar si existe ya el cod. del campo clave primaria
        If ExisteCP(Text1(0)) Then b = False
    End If
    
    
    ' ************************************************************************************
    DatosOk = b
    
EDatosOK:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Function

Private Sub PosicionarData()
Dim cad As String, Indicador As String

    ' *** canviar-ho per tota la PK de la capçalera, no llevar els () ***
    cad = "(codsocio=" & Text1(0).Text & ")"
    
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
    ConseguirFoco Text1(Index), Modo
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

    If Not PerderFocoGnral(Text1(Index), Modo) Then Exit Sub
    
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub
    
    ' ***************** configurar els LostFocus dels camps de la capçalera *****************
    Select Case Index
        Case 0 'cod cliente
            PonerFormatoEntero Text1(0)

        Case 1 'NOMBRE
            Text1(Index).Text = UCase(Text1(Index).Text)
        
        Case 2 'NIF
            Text1(Index).Text = UCase(Text1(Index).Text)
            ValidarNIF Text1(Index).Text
                
        Case 10, 11, 12, 13, 14, 15 'telèfons, fax i mòbils
'            PosarFormatTelefon Text1(Index)
                
        Case 7 'COLECTIVO
            If PonerFormatoEntero(Text1(Index)) Then
                Text2(Index).Text = PonerNombreDeCod(Text1(Index), "scoope", "nomcoope")
                If Text2(Index).Text = "" Then
                    cadMen = "No existe el Colectivo: " & Text1(Index).Text & vbCrLf
                    cadMen = cadMen & "¿Desea crearlo?" & vbCrLf
                    If MsgBox(cadMen, vbQuestion + vbYesNo) = vbYes Then
                        Set frmCoo = New frmManCoope
                        frmCoo.DatosADevolverBusqueda = "0|1|"
                        frmCoo.NuevoCodigo = Text1(Index).Text
                        Text1(Index).Text = ""
                        TerminaBloquear
                        frmCoo.Show vbModal
                        Set frmCoo = Nothing
                        If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Data1, 1
                    Else
                        Text1(Index).Text = ""
                    End If
                    PonerFoco Text1(Index)
                End If
            Else
                Text2(Index).Text = ""
            End If
            
        Case 8 'Situacion
            If PonerFormatoEntero(Text1(Index)) Then
                Text2(Index) = PonerNombreDeCod(Text1(Index), "ssitua", "nomsitua", "codsitua", "N")
                If Text2(Index).Text = "" Then
                    cadMen = "No existe la Situación: " & Text1(Index).Text & vbCrLf
                    cadMen = cadMen & "¿Desea crearla?" & vbCrLf
                    If MsgBox(cadMen, vbQuestion + vbYesNo) = vbYes Then
                        Set frmSit = New frmManSitua
                        frmSit.DatosADevolverBusqueda = "0|1|"
                        frmSit.NuevoCodigo = Text1(Index).Text
                        Text1(Index).Text = ""
                        TerminaBloquear
                        frmSit.Show vbModal
                        Set frmSit = Nothing
                        If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Data1, 1
                    Else
                        Text1(Index).Text = ""
                    End If
                    PonerFoco Text1(Index)
                End If
            Else
                Text2(Index).Text = ""
            End If
                
        Case 17 'Forma de Pago
            If PonerFormatoEntero(Text1(Index)) Then
                Text2(Index) = PonerNombreDeCod(Text1(Index), "sforpa", "nomforpa", "codforpa", "N")
                If Text2(Index).Text = "" Then
                    cadMen = "No existe la F.Pago: " & Text1(Index).Text & vbCrLf
                    cadMen = cadMen & "¿Desea crearla?" & vbCrLf
                    If MsgBox(cadMen, vbQuestion + vbYesNo) = vbYes Then
                        Set frmFpa = New frmManFpago
                        frmFpa.DatosADevolverBusqueda = "0|1|"
                        frmFpa.NuevoCodigo = Text1(Index).Text
                        Text1(Index).Text = ""
                        TerminaBloquear
                        frmFpa.Show vbModal
                        Set frmFpa = Nothing
                        If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Data1, 1
                    Else
                        Text1(Index).Text = ""
                    End If
                    PonerFoco Text1(Index)
                End If
            Else
                Text2(Index).Text = ""
            End If
        
        Case 15, 16 'Fechas
            PonerFormatoFecha Text1(Index)
            
        Case 9 'cuenta contable
            If Text1(Index).Text = "" Then Exit Sub
            If Modo = 3 Then
                Text2(Index).Text = PonerNombreCuenta(Text1(Index), Modo, "") 'text1(0).Text)
            Else
                Text2(Index).Text = PonerNombreCuenta(Text1(Index), Modo, Text1(0).Text)
            End If
            
        Case 23 'DTO. X LITRO
            cadMen = TransformaPuntosComas(Text1(Index).Text)
            Text1(Index).Text = Format(cadMen, "0.0000")
            
        Case 25 'Nro de socio de la cooperativa
            PonerFormatoEntero Text1(Index)
            
    End Select
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
    If KeyAscii = 13 Then 'ENTER
        KeyAscii = 0
        SendKeys "{tab}"
    ElseIf KeyAscii = 27 Then 'ESC
        If (Modo = 0 Or Modo = 2) Then Unload Me
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
Dim sql As String
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
            sql = "¿Seguro que desea eliminar la Tarjeta?"
            sql = sql & vbCrLf & "Tarjeta: " & AdoAux(Index).Recordset!Numtarje
            If MsgBox(sql, vbQuestion + vbYesNo) = vbYes Then
                eliminar = True
                sql = "DELETE FROM starje"
                sql = sql & vWhere & " AND numlinea= " & AdoAux(Index).Recordset!numlinea
            End If
            
        Case 1 'matriculas
            sql = "¿Seguro que desea eliminar la Matricula?"
            sql = sql & vbCrLf & "Nombre: " & AdoAux(Index).Recordset!matricul
            If MsgBox(sql, vbQuestion + vbYesNo) = vbYes Then
                eliminar = True
                sql = "DELETE FROM smatri"
                sql = sql & vWhere & " AND numlinea= " & AdoAux(Index).Recordset!numlinea
            End If
            
            
        Case 2 'margenes
            sql = "¿Seguro que desea eliminar el Margen del Artículo?"
            sql = sql & vbCrLf & "Código: " & AdoAux(Index).Recordset!codartic
            If MsgBox(sql, vbQuestion + vbYesNo) = vbYes Then
                eliminar = True
                sql = "DELETE FROM smargen"
                sql = sql & vWhere & " AND numlinea= " & AdoAux(Index).Recordset!numlinea
            End If
        
            
    End Select

    If eliminar Then
        NumRegElim = AdoAux(Index).Recordset.AbsolutePosition
        TerminaBloquear
        Conn.Execute sql
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
Dim vWhere As String, vtabla As String
Dim anc As Single
Dim i As Integer
    
    ModoLineas = 1 'Posem Modo Afegir Llínia
    
    If (Modo = 3) Or (Modo = 4) Then 'Insertar o Modificar Capçalera
        cmdAceptar_Click
        If ModoLineas = 0 Then Exit Sub
    End If
       
    NumTabMto = Index
    PonerModo 5, Index
    
    ' *** bloquejar la clau primaria de la capçalera ***
    BloquearTxt Text1(0), True

    ' *** posar el nom del les distintes taules de llínies ***
    Select Case Index
        Case 0: vtabla = "starje"
        Case 1: vtabla = "smatri"
        Case 2: vtabla = "smargen"
    End Select
    
    vWhere = ObtenerWhereCab(False)
    
    Select Case Index
        Case 0, 1, 2 ' *** pose els index dels tabs de llínies que tenen datagrid ***
            ' *** canviar la clau primaria de les llínies,
            'pasar a "" si no volem que mos sugerixca res a l'afegir ***
            NumF = SugerirCodigoSiguienteStr(vtabla, "numlinea", vWhere)

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
                    txtAux(0).Text = Text1(0).Text 'codsocio
                    txtAux(3).Text = Text1(1).Text 'nomsocio
                    txtAux(1).Text = NumF 'numlinea
                    txtAux(2).Text = NumTarj
                    For i = 4 To 7
                        txtAux(i).Text = ""
                    Next i
                    txtAux(12).Text = ""
                    cmbAux(0).ListIndex = 1
                    PonerFoco txtAux(2)
                Case 1 'matriculas
                    txtAux(8).Text = Text1(0).Text 'codsocio
                    txtAux(9).Text = NumF 'numlinea
                    For i = 10 To 11
                        txtAux(i).Text = ""
                    Next i
                    
                    PonerFoco txtAux(10)
                Case 2 'margenes
                    txtAux(13).Text = Text1(0).Text 'codsocio
                    txtAux(14).Text = NumF 'numlinea
                    For i = 15 To 17
                        txtAux(i).Text = ""
                    Next i
                    Text2(0).Text = ""
                    PonerFoco txtAux(15)
            End Select
    End Select
End Sub

Private Sub BotonModificarLinea(Index As Integer)
    Dim anc As Single
    Dim i As Integer
    Dim j As Integer
    
    If AdoAux(Index).Recordset.EOF Then Exit Sub
    If AdoAux(Index).Recordset.RecordCount < 1 Then Exit Sub
    
    ModoLineas = 2 'Modificar llínia
       
    If Modo = 4 Then 'Modificar Capçalera
        cmdAceptar_Click
        If ModoLineas = 0 Then Exit Sub
    End If
       
    NumTabMto = Index
    PonerModo 5, Index
    ' *** bloqueje la clau primaria de la capçalera ***
    BloquearTxt Text1(0), True
  
    Select Case Index
        Case 0, 1, 2 ' *** pose els index de llínies que tenen datagrid (en o sense tab) ***
            If DataGridAux(Index).Bookmark < DataGridAux(Index).FirstRow Or DataGridAux(Index).Bookmark > (DataGridAux(Index).FirstRow + DataGridAux(Index).VisibleRows - 1) Then
                i = DataGridAux(Index).Bookmark - DataGridAux(Index).FirstRow
                DataGridAux(Index).Scroll 0, i
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
        
            For j = 0 To 2
                txtAux(j).Text = DataGridAux(Index).Columns(j).Text
            Next j
            
            PosicionarCombo cmbAux(0), AdoAux(Index).Recordset!tiptarje
            
            For j = 3 To 7
                txtAux(j).Text = DataGridAux(Index).Columns(j + 2).Text
            Next j
            txtAux(12).Text = DataGridAux(Index).Columns(10).Text
            For i = 0 To 1
                BloquearTxt txtAux(i), False
            Next i
            
        Case 1 'MATRICULAS
            For j = 8 To 11
                txtAux(j).Text = DataGridAux(Index).Columns(j - 8).Text
            Next j
            
            For i = 8 To 9
                BloquearTxt txtAux(i), False
            Next i
            
        Case 2 'MARGENES
        
            For j = 13 To 14
                txtAux(j).Text = DataGridAux(Index).Columns(j - 13).Text
            Next j
            txtAux(15).Text = DataGridAux(Index).Columns(2).Text
            Text2(0).Text = DataGridAux(Index).Columns(3).Text
            txtAux(16).Text = DataGridAux(Index).Columns(4).Text
            
            For i = 13 To 14
                BloquearTxt txtAux(i), False
            Next i
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
            Text2(0).visible = (xModo = 1)
            Text2(0).Top = alto
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
Dim i As Integer

    ' *** neteje els combos, els pose valor i seleccione el valor per defecte ***
    For i = 0 To 5
        Combo1(i).Clear
    Next i
    cmbAux(0).Clear
    
    For i = 1 To 3
        Combo1(i).AddItem "No"
        Combo1(i).ItemData(Combo1(i).NewIndex) = 0
        Combo1(i).AddItem "Si"
        Combo1(i).ItemData(Combo1(i).NewIndex) = 1
    Next i
    
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
    Combo1(4).AddItem "Cooperativa"
    Combo1(4).ItemData(Combo1(4).NewIndex) = 0
    Combo1(4).AddItem "Tarjetas Visa"
    Combo1(4).ItemData(Combo1(4).NewIndex) = 1
    Combo1(4).AddItem "Crédito local"
    Combo1(4).ItemData(Combo1(4).NewIndex) = 2
    Combo1(4).AddItem "Clientes paso"
    Combo1(4).ItemData(Combo1(4).NewIndex) = 3
    Combo1(4).AddItem "Efectivo"
    Combo1(4).ItemData(Combo1(4).NewIndex) = 4

    ' combo para indicar que el cliente factura con la fp de su ficha independientemente
    ' de la forma de pago del albaran
    Combo1(5).AddItem "No"
    Combo1(5).ItemData(Combo1(5).NewIndex) = 0
    Combo1(5).AddItem "Si"
    Combo1(5).ItemData(Combo1(5).NewIndex) = 1


    ' combo para indicar que al cliente se le envia la factura por email
    Combo1(6).AddItem "No"
    Combo1(6).ItemData(Combo1(6).NewIndex) = 0
    Combo1(6).AddItem "Si"
    Combo1(6).ItemData(Combo1(6).NewIndex) = 1



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
            
        Case 12, 11 'Observaciones dpto
            PonerFocoBtn Me.cmdAceptar
            
        Case 15 ' articulo
            If PonerFormatoEntero(txtAux(Index)) Then
                Text2(0).Text = PonerNombreDeCod(txtAux(Index), "sartic", "nomartic", "codartic", "N")
                If Text2(0).Text = "" Then
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
        
        Case 17 ' euros por kilo
            PonerFormatoDecimal txtAux(Index), 5
            PonerFocoBtn Me.cmdAceptar
            
    End Select
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

Private Function DatosOkLlin(nomFrame As String) As Boolean
Dim Rs As ADODB.Recordset
Dim sql As String
Dim b As Boolean
Dim Cant As Integer
Dim Mens As String
Dim vFact As Byte, vDocum As Byte

    On Error GoTo EDatosOKLlin

    Mens = ""
    DatosOkLlin = False
        
    b = CompForm2(Me, 2, nomFrame) 'Comprovar formato datos ok
    If Not b Then Exit Function
    
    Select Case NumTabMto
        Case 0 ' tarjetas
            ' en el caso de que la tarjeta sea profesional la matricula es obligatoria
            If cmbAux(0).ListIndex = 2 And txtAux(12).Text = "" Then
                MsgBox "Si el tipo de tarjeta es Profesional, debe introducir la matrícula", vbExclamation
                PonerFoco txtAux(12)
                b = False
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
            frmCoo.CodigoActual = Text1(7).Text
            frmCoo.Show vbModal
            Set frmCoo = Nothing
            PonerFoco Text1(7)
            
        Case 1 'situaciones
            Set frmSit = New frmManSitua
            frmSit.DatosADevolverBusqueda = "0|1|"
            frmSit.CodigoActual = Text1(8).Text
            frmSit.Show vbModal
            Set frmSit = Nothing
            PonerFoco Text1(8)
            
        Case 3 'formas de pago
            Set frmFpa = New frmManFpago
            frmFpa.DatosADevolverBusqueda = "0|1|"
            frmFpa.CodigoActual = Text1(17).Text
            frmFpa.Show vbModal
            Set frmFpa = Nothing
            PonerFoco Text1(17)
            
        Case 2 'Cuentas Contables (de contabilidad)
            If vParamAplic.NumeroConta = 0 Then Exit Sub
            
            indice = Index + 7
            Set frmCtas = New frmCtasConta
            frmCtas.NumDigit = 0
            frmCtas.DatosADevolverBusqueda = "0|1|"
            frmCtas.CodigoActual = txtAux(indice).Text
            frmCtas.Show vbModal
            Set frmCtas = Nothing
            PonerFoco Text1(indice)
            
    End Select
    
    If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Data1, 1
End Sub

Private Sub frmCtas_DatoSeleccionado(CadenaSeleccion As String)
'Cuentas contables de la Contabilidad
    Text1(indice).Text = RecuperaValor(CadenaSeleccion, 1) 'codmacta
    Text2(indice).Text = RecuperaValor(CadenaSeleccion, 2) 'des macta
End Sub

Private Sub frmCoo_DatoSeleccionado(CadenaSeleccion As String)
'Mantenimiento de Colectivos
    Text1(7).Text = RecuperaValor(CadenaSeleccion, 1) 'codcoope
    FormateaCampo Text1(7)
    Text2(7).Text = RecuperaValor(CadenaSeleccion, 2) 'nomcoope
End Sub

Private Sub frmSit_DatoSeleccionado(CadenaSeleccion As String)
'Mantenimiento Grupo
    Text1(8).Text = RecuperaValor(CadenaSeleccion, 1) 'codsitua
    FormateaCampo Text1(8)
    Text2(8).Text = RecuperaValor(CadenaSeleccion, 2) 'nomsitua
End Sub

Private Sub frmFpa_DatoSeleccionado(CadenaSeleccion As String)
'Mantenimiento F.Pago
    Text1(17).Text = RecuperaValor(CadenaSeleccion, 1) 'codforpa
    FormateaCampo Text1(17)
    Text2(17).Text = RecuperaValor(CadenaSeleccion, 2) 'nomforpa
End Sub


Private Sub imgWeb_Click(Index As Integer)
    'Abrimos el explorador de windows con la pagina Web del cliente
    If Modo = 0 Then Exit Sub
    Screen.MousePointer = vbHourglass
'    If LanzaHome("websoporte") Then espera 2
    If LanzaHomeGnral(Text1(10).Text) Then espera 2
    Screen.MousePointer = vbDefault
End Sub

Private Sub DataGridAux_RowColChange(Index As Integer, LastRow As Variant, ByVal LastCol As Integer)
Dim i As Byte

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
                For i = 21 To 24
'                   txtAux(i).Text = ""
                Next i
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
Dim i As Byte

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
Dim i As Byte
Dim tots As String

    On Error GoTo ECarga

    tots = MontaSQLCarga(Index, enlaza)

    CargaGridGnral Me.DataGridAux(Index), Me.AdoAux(Index), tots, PrimeraVez
    
    Select Case Index
        Case 0 'tarjetas
            'si es visible|control|tipo campo|nombre campo|ancho control|
            tots = "N||||0|;N||||0|;" 'codsocio,numlinea
            tots = tots & "S|txtAux(2)|T|Tarjeta|1000|;N||||0|;S|cmbAux(0)|C|Tipo|1200|;"
            tots = tots & "S|txtAux(3)|T|Nombre|3500|;S|txtAux(4)|T|Banco|650|;"
            tots = tots & "S|txtAux(5)|T|Sucur.|650|;S|txtAux(6)|T|DC|380|;"
            tots = tots & "S|txtAux(7)|T|Cuenta|1500|;"
            tots = tots & "S|txtAux(12)|T|Matricula|1500|;"
            
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
Dim nomFrame As String
Dim b As Boolean

    On Error Resume Next

    ' *** posa els noms del frames, tant si son de grid com si no ***
    Select Case NumTabMto
        Case 0: nomFrame = "FrameAux0" 'tarjetas
        Case 1: nomFrame = "FrameAux1" 'matriculas
        Case 2: nomFrame = "FrameAux2" 'margenes
    End Select
    
    If DatosOkLlin(nomFrame) Then
        TerminaBloquear
        If InsertarDesdeForm2(Me, 2, nomFrame) Then
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
Dim nomFrame As String
Dim V As Integer
    
    On Error Resume Next

    ' *** posa els noms del frames, tant si son de grid com si no ***
    Select Case NumTabMto
        Case 0: nomFrame = "FrameAux0" 'tarjetas
        Case 1: nomFrame = "FrameAux1" 'Matriculas
        Case 2: nomFrame = "FrameAux2" 'Margenes
    End Select
    ModificarLinea = False
    If DatosOkLlin(nomFrame) Then
        TerminaBloquear
        If ModificaDesdeFormulario2(Me, 2, nomFrame) Then
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
    vWhere = vWhere & " codsocio=" & Val(Text1(0).Text)
    
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

