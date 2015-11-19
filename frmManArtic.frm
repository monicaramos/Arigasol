VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmManArtic 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Artículos"
   ClientHeight    =   7455
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11715
   Icon            =   "frmManArtic.frx":0000
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
      TabIndex        =   44
      Top             =   480
      Width           =   11295
      Begin VB.TextBox text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   0
         Left            =   1440
         MaxLength       =   6
         TabIndex        =   0
         Tag             =   "Código de articulo|N|N|0|999999|sartic|codartic|000000|S|"
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox text1 
         Height          =   285
         Index           =   1
         Left            =   3480
         MaxLength       =   40
         TabIndex        =   1
         Tag             =   "Nombre|T|N|||sartic|nomartic|||"
         Top             =   240
         Width           =   4140
      End
      Begin VB.Label Label4 
         Caption         =   "Nombre "
         Height          =   255
         Left            =   2640
         TabIndex        =   46
         Top             =   240
         Width           =   660
      End
      Begin VB.Label Label1 
         Caption         =   "Código"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   45
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      Height          =   555
      Index           =   0
      Left            =   240
      TabIndex        =   41
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
         TabIndex        =   42
         Top             =   180
         Width           =   2655
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   10410
      TabIndex        =   30
      Top             =   6960
      Width           =   1035
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   9120
      TabIndex        =   29
      Top             =   6960
      Width           =   1035
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5460
      Left            =   270
      TabIndex        =   43
      Top             =   1350
      Width           =   11295
      _ExtentX        =   19923
      _ExtentY        =   9631
      _Version        =   393216
      Style           =   1
      TabsPerRow      =   6
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
      TabPicture(0)   =   "frmManArtic.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label5"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label19"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "imgBuscar(0)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "imgBuscar(1)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label8"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label2"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label17"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "imgBuscar(2)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "imgBuscar(3)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "imgBuscar(5)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label20"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "text1(2)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "text2(3)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "text1(3)"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "FrameDatosAlta"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "FrameDatosContacto"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "text1(4)"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "text2(4)"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "text1(5)"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "Frame3"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "Frame4"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "text2(17)"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "text1(17)"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "text2(0)"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "text1(20)"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "text2(20)"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).ControlCount=   26
      TabCaption(1)   =   "Tarifas Venta"
      TabPicture(1)   =   "frmManArtic.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "FrameAux0"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Bonificaciones"
      TabPicture(2)   =   "frmManArtic.frx":0044
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "FrameAux1"
      Tab(2).ControlCount=   1
      Begin VB.TextBox text2 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   20
         Left            =   2700
         TabIndex        =   84
         Top             =   2490
         Width           =   2895
      End
      Begin VB.TextBox text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   20
         Left            =   1440
         MaxLength       =   10
         TabIndex        =   7
         Tag             =   "Cta.Compras|T|S|||sartic|ctacompr|||"
         Top             =   2490
         Width           =   1215
      End
      Begin VB.TextBox text2 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   2010
         TabIndex        =   79
         Top             =   2070
         Width           =   3585
      End
      Begin VB.TextBox text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   17
         Left            =   1440
         MaxLength       =   10
         TabIndex        =   5
         Tag             =   "Cta.Cont.Cliente|T|S|||sartic|codmaccl|||"
         Top             =   1710
         Width           =   1215
      End
      Begin VB.TextBox text2 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   17
         Left            =   2700
         TabIndex        =   77
         Top             =   1710
         Width           =   2895
      End
      Begin VB.Frame Frame4 
         Caption         =   "Datos para Combustibles"
         ForeColor       =   &H00972E0B&
         Height          =   1965
         Left            =   5880
         TabIndex        =   72
         Top             =   2970
         Width           =   5295
         Begin VB.TextBox text1 
            Height          =   285
            Index           =   23
            Left            =   3780
            MaxLength       =   15
            TabIndex        =   28
            Tag             =   "Código Externo|T|S|||sartic|codexterno|||"
            Text            =   "000000000000000"
            Top             =   1530
            Width           =   1455
         End
         Begin VB.CheckBox chkDomicilio 
            Caption         =   "A Domicilio"
            Height          =   315
            Index           =   0
            Left            =   180
            TabIndex        =   27
            Tag             =   "A Domicilio|N|N|0|1|sartic|esdomiciliado||N|"
            Top             =   1500
            Width           =   1215
         End
         Begin VB.TextBox text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   18
            Left            =   3780
            MaxLength       =   7
            TabIndex        =   24
            Tag             =   "Porcentaje Biodiesel|N|N|0.00|999.99|sartic|porcbd|##0.00||"
            Top             =   720
            Width           =   1125
         End
         Begin VB.CheckBox chkAux 
            BackColor       =   &H80000005&
            Height          =   195
            Index           =   0
            Left            =   3780
            TabIndex        =   22
            Tag             =   "Declara GP|N|S|||sartic|gp|||"
            Top             =   390
            Width           =   165
         End
         Begin VB.TextBox text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   16
            Left            =   3780
            MaxLength       =   6
            TabIndex        =   26
            Tag             =   "Impuesto|N|S|0|9.9999|sartic|impuesto|0.0000||"
            Top             =   1140
            Width           =   1125
         End
         Begin VB.ComboBox Combo1 
            Height          =   315
            Index           =   0
            Left            =   1200
            Style           =   2  'Dropdown List
            TabIndex        =   25
            Tag             =   "Tipo Gasolina|N|N|0|4|sartic|tipogaso|0|N|"
            Top             =   1140
            Width           =   1335
         End
         Begin VB.TextBox text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   14
            Left            =   1200
            MaxLength       =   3
            TabIndex        =   21
            Tag             =   "Tanque|N|S|0|999|sartic|numtanqu|000||"
            Top             =   360
            Width           =   600
         End
         Begin VB.TextBox text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   15
            Left            =   1200
            MaxLength       =   3
            TabIndex        =   23
            Tag             =   "Manguera|N|S|0|999|sartic|nummangu|000||"
            Top             =   750
            Width           =   600
         End
         Begin VB.Label Label1 
            Caption         =   "Código Externo"
            Height          =   255
            Index           =   3
            Left            =   2640
            TabIndex        =   88
            Top             =   1560
            Width           =   1230
         End
         Begin VB.Label Label1 
            Caption         =   "Porc.Biodiesel"
            Height          =   255
            Index           =   2
            Left            =   2640
            TabIndex        =   81
            Top             =   750
            Width           =   1005
         End
         Begin VB.Label Label1 
            Caption         =   "Declara GP"
            Height          =   255
            Index           =   1
            Left            =   2640
            TabIndex        =   80
            Top             =   390
            Width           =   855
         End
         Begin VB.Label Label15 
            Caption         =   "Impuesto"
            Height          =   255
            Left            =   2655
            TabIndex        =   76
            Top             =   1140
            Width           =   975
         End
         Begin VB.Label Label13 
            Caption         =   "Tipo Gasol."
            Height          =   255
            Left            =   150
            TabIndex        =   75
            Top             =   1170
            Width           =   975
         End
         Begin VB.Label Label12 
            Caption         =   "Nº Tanque"
            Height          =   255
            Left            =   150
            TabIndex        =   74
            Top             =   390
            Width           =   855
         End
         Begin VB.Label Label7 
            Caption         =   "Nº Manguera"
            Height          =   255
            Left            =   150
            TabIndex        =   73
            Top             =   765
            Width           =   975
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Datos Facturación"
         ForeColor       =   &H00972E0B&
         Height          =   1950
         Left            =   150
         TabIndex        =   69
         Top             =   2970
         Width           =   5535
         Begin VB.TextBox text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   21
            Left            =   1320
            MaxLength       =   14
            TabIndex        =   11
            Tag             =   "Peso Artículo|N|S|||sartic|pesoart|#,##0.00||"
            Top             =   1350
            Width           =   1320
         End
         Begin VB.TextBox text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   22
            Left            =   4140
            MaxLength       =   6
            TabIndex        =   12
            Tag             =   "Precio Sigaus|N|S|||sartic|precsigaus|#,###0.0000||"
            Top             =   1350
            Width           =   1125
         End
         Begin VB.TextBox text2 
            BackColor       =   &H80000018&
            Enabled         =   0   'False
            Height          =   285
            Index           =   19
            Left            =   2295
            TabIndex        =   82
            Top             =   840
            Width           =   2985
         End
         Begin VB.TextBox text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   19
            Left            =   1320
            MaxLength       =   6
            TabIndex        =   10
            Tag             =   "Art.Dto.|N|S|0|999999|sartic|artdto|000000||"
            Top             =   840
            Width           =   900
         End
         Begin VB.TextBox text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   6
            Left            =   1320
            MaxLength       =   14
            TabIndex        =   8
            Tag             =   "P.V.P.|N|N|0|99999999|sartic|preventa|#,###,##0.000||"
            Top             =   360
            Width           =   1200
         End
         Begin VB.TextBox text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   7
            Left            =   4080
            MaxLength       =   6
            TabIndex        =   9
            Tag             =   "Bonificación|N|N|0|9.9999|sartic|bonigral|0.0000||"
            Top             =   360
            Width           =   1215
         End
         Begin VB.Label Label22 
            Caption         =   "Peso Unidad"
            Height          =   255
            Left            =   120
            TabIndex        =   87
            Top             =   1365
            Width           =   975
         End
         Begin VB.Label Label23 
            Caption         =   "Precio Sigaus"
            Height          =   255
            Left            =   2910
            TabIndex        =   86
            Top             =   1380
            Width           =   975
         End
         Begin VB.Label Label18 
            Caption         =   "Art.Dto."
            Height          =   255
            Left            =   135
            TabIndex        =   83
            Top             =   840
            Width           =   705
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   4
            Left            =   1035
            ToolTipText     =   "Buscar Artículo"
            Top             =   840
            Width           =   240
         End
         Begin VB.Label Label6 
            Caption         =   "P.V.P."
            Height          =   255
            Left            =   120
            TabIndex        =   71
            Top             =   375
            Width           =   615
         End
         Begin VB.Label Label3 
            Caption         =   "Bonif.Gral."
            Height          =   255
            Left            =   2880
            TabIndex        =   70
            Top             =   360
            Width           =   975
         End
      End
      Begin VB.TextBox text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   5
         Left            =   1440
         MaxLength       =   2
         TabIndex        =   6
         Tag             =   "Tipo IVA|N|N|0|99|sartic|codigiva|||"
         Top             =   2070
         Width           =   495
      End
      Begin VB.Frame FrameAux1 
         BorderStyle     =   0  'None
         Height          =   4215
         Left            =   -74760
         TabIndex        =   65
         Top             =   480
         Width           =   10695
         Begin VB.TextBox txtAux 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   290
            Index           =   12
            Left            =   4440
            MaxLength       =   30
            TabIndex        =   36
            Tag             =   "Bonificacion|N|N|0|999.999|sbonif|bonifica|###,##0.000||"
            Top             =   3600
            Visible         =   0   'False
            Width           =   1035
         End
         Begin VB.ComboBox cmbAux 
            Height          =   315
            Index           =   0
            Left            =   1320
            TabIndex        =   33
            Tag             =   "Tipo Cliente|N|N|0|2|sbonif|tipsocio|0|S|"
            Text            =   "Combo2"
            Top             =   3600
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.TextBox txtAux 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   290
            Index           =   8
            Left            =   0
            MaxLength       =   6
            TabIndex        =   31
            Tag             =   "Código de Articulo|N|N|1|999999|sbonif|codartic|000000|S|"
            Top             =   3600
            Visible         =   0   'False
            Width           =   555
         End
         Begin VB.TextBox txtAux 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   290
            Index           =   10
            Left            =   2280
            MaxLength       =   10
            TabIndex        =   34
            Tag             =   "Desde cantidad|N|N|0|9999999999|sbonif|desdecan|#,###,###,##0||"
            Top             =   3600
            Visible         =   0   'False
            Width           =   795
         End
         Begin VB.TextBox txtAux 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   290
            Index           =   11
            Left            =   3240
            MaxLength       =   30
            TabIndex        =   35
            Tag             =   "Hasta cantidad|N|N|0|9999999999|sbonif|hastacan|#,###,###,###||"
            Top             =   3600
            Visible         =   0   'False
            Width           =   1035
         End
         Begin VB.TextBox txtAux 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   290
            Index           =   9
            Left            =   600
            MaxLength       =   2
            TabIndex        =   32
            Tag             =   "Numero linea|N|N|1|99|sbonif|numlinea|00|S|"
            Text            =   "linea"
            Top             =   3600
            Visible         =   0   'False
            Width           =   555
         End
         Begin MSComctlLib.Toolbar ToolAux 
            Height          =   390
            Index           =   1
            Left            =   0
            TabIndex        =   66
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
            Bindings        =   "frmManArtic.frx":0060
            Height          =   3645
            Index           =   1
            Left            =   0
            TabIndex        =   67
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
         Height          =   4215
         Left            =   -74760
         TabIndex        =   62
         Top             =   480
         Width           =   10695
         Begin VB.TextBox txtAux 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   290
            Index           =   0
            Left            =   -120
            MaxLength       =   6
            TabIndex        =   37
            Tag             =   "Código de articulo|N|N|1|999999|starif|codartic|000000|S|"
            Top             =   3720
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.TextBox txtAux 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   290
            Index           =   1
            Left            =   960
            MaxLength       =   1
            TabIndex        =   38
            Tag             =   "Tarifa|N|N|0|9|starif|codtarif|0|S|"
            Top             =   3720
            Visible         =   0   'False
            Width           =   1275
         End
         Begin VB.TextBox txtAux 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   290
            Index           =   2
            Left            =   2400
            MaxLength       =   14
            TabIndex        =   39
            Tag             =   "Precio Venta|N|N|0||starif|preventa|#,###,##0.000||"
            Top             =   3720
            Visible         =   0   'False
            Width           =   1455
         End
         Begin MSComctlLib.Toolbar ToolAux 
            Height          =   390
            Index           =   0
            Left            =   0
            TabIndex        =   63
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
            Bindings        =   "frmManArtic.frx":0078
            Height          =   3645
            Index           =   0
            Left            =   0
            TabIndex        =   64
            Top             =   480
            Width           =   9675
            _ExtentX        =   17066
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
      End
      Begin VB.TextBox text2 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   4
         Left            =   2700
         TabIndex        =   59
         Top             =   1320
         Width           =   2895
      End
      Begin VB.TextBox text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   4
         Left            =   1440
         MaxLength       =   10
         TabIndex        =   4
         Tag             =   "Cta.Contable|T|S|||sartic|codmacta|||"
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Frame FrameDatosContacto 
         Caption         =   "Datos Almacen"
         ForeColor       =   &H00972E0B&
         Height          =   1515
         Left            =   5895
         TabIndex        =   52
         Top             =   570
         Width           =   5310
         Begin VB.CheckBox chkNuevo 
            Caption         =   "Creado Automáticamente"
            Height          =   315
            Index           =   0
            Left            =   2790
            TabIndex        =   18
            Tag             =   "Artículo Nuevo|N|N|0|1|sartic|artnuevo||N|"
            Top             =   1080
            Width           =   2205
         End
         Begin VB.CheckBox chkCtrStock 
            Caption         =   "¿Control de stock?"
            Height          =   315
            Index           =   0
            Left            =   240
            TabIndex        =   17
            Tag             =   "Control de stock|N|N|0|1|sartic|ctrstock||N|"
            Top             =   1080
            Width           =   1815
         End
         Begin VB.TextBox text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   8
            Left            =   1215
            MaxLength       =   14
            TabIndex        =   13
            Tag             =   "Stock|N|N|||sartic|canstock|#,###,##0.000||"
            Top             =   360
            Width           =   1215
         End
         Begin VB.TextBox text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   10
            Left            =   1215
            MaxLength       =   14
            TabIndex        =   15
            Tag             =   "P.M.P.|N|S|0|9999999.999|sartic|preciopmp|#,###,##0.00000||"
            Top             =   730
            Width           =   1215
         End
         Begin VB.TextBox text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   9
            Left            =   3825
            MaxLength       =   14
            TabIndex        =   14
            Tag             =   "U.Precio|N|S|0|9999999.999|sartic|ultpreci|#,###,##0.00000||"
            Top             =   360
            Width           =   1215
         End
         Begin VB.TextBox text1 
            Height          =   285
            Index           =   11
            Left            =   3825
            MaxLength       =   10
            TabIndex        =   16
            Tag             =   "F.U.Compra|F|S|||sartic|ultfecha|dd/mm/yyyy||"
            Top             =   720
            Width           =   1215
         End
         Begin VB.Image imgFec 
            Height          =   240
            Index           =   11
            Left            =   3480
            Picture         =   "frmManArtic.frx":0090
            ToolTipText     =   "Buscar fecha"
            Top             =   720
            Width           =   240
         End
         Begin VB.Label Label9 
            Caption         =   "Stock"
            Height          =   255
            Left            =   240
            TabIndex        =   58
            Top             =   360
            Width           =   495
         End
         Begin VB.Label Label10 
            Caption         =   "Precio M.P."
            Height          =   255
            Left            =   240
            TabIndex        =   55
            Top             =   735
            Width           =   855
         End
         Begin VB.Label Label14 
            Caption         =   "U. Precio"
            Height          =   255
            Left            =   2790
            TabIndex        =   54
            Top             =   375
            Width           =   735
         End
         Begin VB.Label Label16 
            Caption         =   "Fecha"
            Height          =   255
            Left            =   2790
            TabIndex        =   53
            Top             =   750
            Width           =   495
         End
      End
      Begin VB.Frame FrameDatosAlta 
         Caption         =   "Datos Inventario"
         ForeColor       =   &H00972E0B&
         Height          =   780
         Left            =   5880
         TabIndex        =   49
         Top             =   2130
         Width           =   5325
         Begin VB.TextBox text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   13
            Left            =   3840
            MaxLength       =   14
            TabIndex        =   20
            Tag             =   "Stock Inv.|N|N|0|9999999.999|sartic|stockinv|#,###,##0.000||"
            Top             =   345
            Width           =   1275
         End
         Begin VB.TextBox text1 
            Height          =   285
            Index           =   12
            Left            =   1200
            MaxLength       =   10
            TabIndex        =   19
            Tag             =   "F.Inv.|F|S|||sartic|fechainv|dd/mm/yyyy||"
            Top             =   345
            Width           =   1200
         End
         Begin VB.Image imgFec 
            Height          =   240
            Index           =   12
            Left            =   930
            Picture         =   "frmManArtic.frx":011B
            ToolTipText     =   "Buscar fecha"
            Top             =   345
            Width           =   240
         End
         Begin VB.Label Label11 
            Caption         =   "Stock Inv."
            Height          =   255
            Left            =   2790
            TabIndex        =   61
            Top             =   345
            Width           =   975
         End
         Begin VB.Label Label21 
            Caption         =   "Fecha"
            Height          =   255
            Left            =   135
            TabIndex        =   51
            Top             =   360
            Width           =   615
         End
      End
      Begin VB.TextBox text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   3
         Left            =   1440
         MaxLength       =   3
         TabIndex        =   3
         Tag             =   "Familia|N|N|0|999|sartic|codfamia|000||"
         Top             =   920
         Width           =   495
      End
      Begin VB.TextBox text2 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   3
         Left            =   1980
         TabIndex        =   40
         Top             =   920
         Width           =   3615
      End
      Begin VB.TextBox text1 
         Height          =   285
         Index           =   2
         Left            =   1440
         MaxLength       =   13
         TabIndex        =   2
         Tag             =   "EAN|T|S|||sartic|codigean|||"
         Top             =   520
         Width           =   1200
      End
      Begin VB.Label Label20 
         Caption         =   "Cta.Compras"
         Height          =   255
         Left            =   210
         TabIndex        =   85
         Top             =   2490
         Width           =   915
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   5
         Left            =   1170
         ToolTipText     =   "Buscar Cta.Contable"
         Top             =   2490
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   3
         Left            =   1170
         ToolTipText     =   "Buscar Iva"
         Top             =   2070
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   2
         Left            =   1170
         ToolTipText     =   "Buscar Cta.Contable"
         Top             =   1710
         Width           =   240
      End
      Begin VB.Label Label17 
         Caption         =   "Cta.Cliente"
         Height          =   255
         Left            =   210
         TabIndex        =   78
         Top             =   1710
         Width           =   795
      End
      Begin VB.Label Label2 
         Caption         =   "Tipo IVA"
         Height          =   255
         Left            =   210
         TabIndex        =   68
         Top             =   2070
         Width           =   735
      End
      Begin VB.Label Label8 
         Caption         =   "Cta.Socio"
         Height          =   255
         Left            =   210
         TabIndex        =   60
         Top             =   1320
         Width           =   735
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   1
         Left            =   1170
         ToolTipText     =   "Buscar Cta.Contable"
         Top             =   1320
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   0
         Left            =   1170
         ToolTipText     =   "Buscar Familia"
         Top             =   915
         Width           =   240
      End
      Begin VB.Label Label19 
         Caption         =   "Familia"
         Height          =   255
         Left            =   210
         TabIndex        =   48
         Top             =   915
         Width           =   855
      End
      Begin VB.Label Label5 
         Caption         =   "Código EAN"
         Height          =   255
         Left            =   210
         TabIndex        =   47
         Top             =   525
         Width           =   1095
      End
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   330
      Left            =   4200
      Top             =   6960
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
      TabIndex        =   56
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
            Style           =   3
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
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
            Object.ToolTipText     =   "Primero"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Anterior"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Siguiente"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
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
         TabIndex        =   57
         Top             =   120
         Width           =   1215
      End
   End
   Begin VB.CommandButton cmdRegresar 
      Caption         =   "&Regresar"
      Height          =   375
      Left            =   10440
      TabIndex        =   50
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
Attribute VB_Name = "frmManArtic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' +-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+
' +-+- Autor: MANOLO                   -+-+
' +-+- Menú: ARTICULOS                 -+-+
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

' *** declarar els formularis als que vaig a cridar ***
Private WithEvents frmB As frmBuscaGrid
Attribute frmB.VB_VarHelpID = -1
Private WithEvents frmB1 As frmBuscaGrid
Attribute frmB1.VB_VarHelpID = -1
Private WithEvents frmC As frmCal 'calendario fecha
Attribute frmC.VB_VarHelpID = -1

Private WithEvents frmFam As frmManFamia 'Familias
Attribute frmFam.VB_VarHelpID = -1

Private WithEvents frmCtas As frmCtasConta 'cuentas contables
Attribute frmCtas.VB_VarHelpID = -1
'VRS:4.0.1
Private WithEvents frmTipIva As frmTipIVAConta  'Tipos de IVA de la contabilidad
Attribute frmTipIva.VB_VarHelpID = -1
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


Private Sub chkAux_Click(Index As Integer)
    If Modo = 1 Then
        'Busqueda
        If InStr(1, BuscaChekc, "chkAux(" & Index & ")") = 0 Then BuscaChekc = BuscaChekc & "chkAux(" & Index & ")|"
    End If
End Sub

Private Sub chkAux_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub chkctrstock_Click(Index As Integer)
    If Modo = 1 Then
        'Busqueda
        If InStr(1, BuscaChekc, "chkctrstock(" & Index & ")") = 0 Then BuscaChekc = BuscaChekc & "chkctrstock(" & Index & ")|"
    End If
End Sub

Private Sub chkctrstock_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub chkDomicilio_Click(Index As Integer)
    If Modo = 1 Then
        'Busqueda
        If InStr(1, BuscaChekc, "chkDomicilio(" & Index & ")") = 0 Then BuscaChekc = BuscaChekc & "chkDomicilio(" & Index & ")|"
    End If

End Sub

Private Sub chkDomicilio_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub chknuevo_Click(Index As Integer)
    If Modo = 1 Then
        'Busqueda
        If InStr(1, BuscaChekc, "chknuevo(" & Index & ")") = 0 Then BuscaChekc = BuscaChekc & "chknuevo(" & Index & ")|"
    End If
End Sub

Private Sub chkNuevo_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
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
                    ModificarLinea
                    PosicionarData
            End Select
        ' **************************
    End Select
    Screen.MousePointer = vbDefault
 
Error1:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub

Private Sub Combo1_GotFocus(Index As Integer)
' *** si n'hi han combos a la capçalera ***
    If Modo = 1 Then Combo1(Index).BackColor = vbYellow
End Sub

Private Sub Combo1_LostFocus(Index As Integer)
    If Combo1(0).BackColor = vbYellow Then Combo1(0).BackColor = vbWhite
    
    Me.chkDomicilio(0).visible = (Combo1(0).ListIndex = 3)
    Me.chkDomicilio(0).Enabled = (Combo1(0).ListIndex = 3)
    If Not chkDomicilio(0).visible Then
        chkDomicilio(0).Value = 0
    End If
   
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
    btnPrimero = 16 'index del botó "primero"
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
        'el 10 i el 11 son separadors
        .Buttons(12).Image = 10  'Imprimir
        .Buttons(13).Image = 11  'Eixir
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
    Next i
    
    'cargar IMAGES de busqueda
    For i = 0 To imgBuscar.Count - 1
        Me.imgBuscar(i).Picture = frmPpal.imgListImages16.ListImages(1).Picture
    Next i
    
    ' *** si n'hi han tabs, per a que per defecte sempre es pose al 1r***
    Me.SSTab1.Tab = 0
    ' *******************************************************************
    
    LimpiarCampos   'Neteja els camps TextBox
    ' ******* si n'hi han llínies *******
    DataGridAux(0).ClearFields
    DataGridAux(1).ClearFields
    
    '*** canviar el nom de la taula i l'ordenació de la capçalera ***
    NombreTabla = "sartic"
    Ordenacion = " ORDER BY codartic"
    
    'Mirem com està guardat el valor del check
    chkVistaPrevia(0).Value = CheckValueLeer(Name)
    
    Data1.ConnectionString = Conn
    '***** cambiar el nombre de la PK de la cabecera *************
    Data1.RecordSource = "Select * from " & NombreTabla & " where codartic=-1"
    Data1.Refresh
       
    ModoLineas = 0
    CargaCombo 0
    
    If DatosADevolverBusqueda = "" Then
        PonerModo 0
    Else
        PonerModo 1 'búsqueda
        ' *** posar de groc els camps visibles de la clau primaria de la capçalera ***
        Text1(0).BackColor = vbYellow 'codartic
        ' ****************************************************************************
    End If
End Sub

Private Sub LimpiarCampos()
    On Error Resume Next
    
    Limpiar Me   'Mètode general: Neteja els controls TextBox
    lblIndicador.Caption = ""
    
    Me.Combo1(0).ListIndex = -1
    Me.cmbAux(0).ListIndex = -1
    Me.chkAux(0).Value = 0
    Me.chkCtrStock(0).Value = 0
    Me.chkNuevo(0).Value = 0
    Me.chkDomicilio(0).Value = 0
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
    BloquearCombo Me, Modo
    BloquearChk chkCtrStock(0), Not b
    BloquearChk chkAux(0), Not b
    BloquearChk chkNuevo(0), Not b
    BloquearChk chkDomicilio(0), Not b
    
    Me.chkDomicilio(0).visible = ((Combo1(0).ListIndex = 3)) Or Modo = 1
    Me.chkDomicilio(0).Enabled = ((Combo1(0).ListIndex = 3) And Modo <> 0 And Modo <> 2) Or Modo = 1
    
    ' ***** bloquejar tots els controls visibles de la clau primaria de la capçalera ***
'    If Modo = 4 Then _
'        BloquearTxt Text1(0), True 'si estic en  modificar, bloqueja la clau primaria
    ' **********************************************************************************
    
    ' **** si n'hi han imagens de buscar en la capçalera *****
    BloquearImgBuscar Me, Modo, ModoLineas
    BloquearImgFec Me, 11, Modo
    BloquearImgFec Me, 12, Modo
    
    chkVistaPrevia(0).Enabled = (Modo <= 2)
    
    PonerLongCampos

    If (Modo < 2) Or (Modo = 3) Then
        CargaGrid 0, False
        CargaGrid 1, False
    End If
    
    b = (Modo = 4) Or (Modo = 2)
    DataGridAux(0).Enabled = b
    DataGridAux(1).Enabled = b
      
    ' ****** si n'hi han combos a la capçalera ***********************
     If (Modo = 0) Or (Modo = 2) Or (Modo = 5) Then
        Combo1(0).Enabled = False
        Combo1(0).BackColor = &H80000018 'groc
    ElseIf (Modo = 1) Or (Modo = 3) Or (Modo = 4) Then
        Combo1(0).Enabled = True
        Combo1(0).BackColor = &H80000005 'blanc
    End If
    ' ****************************************************************
    
    PonerModoOpcionesMenu (Modo) 'Activar opcions menú según modo
    PonerOpcionesMenu   'Activar opcions de menú según nivell
                        'de permisos de l'usuari
    
    ' ### [Monica] 04/09/2006
    'bloqueamos el frame de combustible de los articulos que no sean combustible
    BloquearDatosCombustible (Modo)

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
    Me.mnNuevo.Enabled = b
    
    b = (Modo = 2 And Data1.Recordset.RecordCount > 0) And Not DeConsulta
    'Modificar
    Toolbar1.Buttons(8).Enabled = b
    Me.mnModificar.Enabled = b
    'eliminar
    Toolbar1.Buttons(9).Enabled = b
    Me.mnEliminar.Enabled = b
    
    'Imprimir
    'Toolbar1.Buttons(12).Enabled = (b Or Modo = 0)
    Toolbar1.Buttons(12).Enabled = b
    Me.mnImprimir.Enabled = b
       
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
Dim SQL As String
Dim tabla As String
    
    ' ********* si n'hi han tabs, dona igual si en datagrid o no ***********
    Select Case Index
               
        Case 0 'TARIFAS
            SQL = "SELECT codartic,codtarif,preventa "
            SQL = SQL & " FROM starif"
            If enlaza Then
                SQL = SQL & ObtenerWhereCab(True)
            Else
                SQL = SQL & " WHERE starif.codartic = -1"
            End If
            SQL = SQL & " ORDER BY starif.codtarif"
            
        Case 1 'BONIFICACIONES
            SQL = "SELECT codartic,numlinea,tipsocio,CASE tipsocio WHEN 0 THEN ""Particular"" WHEN 1 THEN ""Profesional"" WHEN 2 THEN ""Comercio"" END, desdecan,hastacan,bonifica "
            SQL = SQL & " FROM sbonif "
            If enlaza Then
                SQL = SQL & ObtenerWhereCab(True)
            Else
                SQL = SQL & " WHERE sbonif.codartic = -1"
            End If
            SQL = SQL & " ORDER BY sbonif.tipsocio"
               
    End Select
    
    MontaSQLCarga = SQL
End Function

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

Private Sub frmB1_Selecionado(CadenaDevuelta As String)
    Dim Aux As String
    
    If CadenaDevuelta <> "" Then
        HaDevueltoDatos = True
        Text1(19).Text = RecuperaValor(CadenaDevuelta, 1)
        text2(19).Text = RecuperaValor(CadenaDevuelta, 2)
    End If

End Sub

'VRS:4.0.1
Private Sub frmTipIva_DatoSeleccionado(CadenaSeleccion As String)
'Tipos de IVA (de la Contabilidad)
    Text1(5).Text = RecuperaValor(CadenaSeleccion, 1) 'codigiva
    FormateaCampo Text1(5)
    text2(0).Text = RecuperaValor(CadenaSeleccion, 2) '% iva

End Sub

' *** si n'hi ha buscar data, posar a les <=== el menor index de les imagens de buscar data ***
' NOTA: ha de coincidir l'index de la image en el del camp a on va a parar el valor
Private Sub imgFec_Click(Index As Integer)
    Dim esq As Long
    Dim dalt As Long
    Dim menu As Long
    Dim obj As Object

    Set frmC = New frmCal
    
    ' *** adrede ***
'    If Index <> 49 Then
        esq = imgFec(Index).Left
        dalt = imgFec(Index).Top
'   Else
'       esq = btnFec(Index).Left
'       dalt = btnFec(Index).Top
'   End If
    
    ' *** adrede ***
'    If Index <> 49 Then
        Set obj = imgFec(Index).Container

        While imgFec(Index).Parent.Name <> obj.Name
            esq = esq + obj.Left
            dalt = dalt + obj.Top
            Set obj = obj.Container
        Wend
'    Else
'        Set obj = btnFec(Index).Container
'
'        While btnFec(Index).Parent.Name <> obj.Name
'            esq = esq + obj.Left
'            dalt = dalt + obj.Top
'            Set obj = obj.Container
'        Wend
'
'    End If
    ' *************
    
    menu = Me.Height - Me.ScaleHeight 'ací tinc el heigth del menú i de la toolbar

    ' *** adredre ***
'    If Index <> 49 Then 'dreta i baix
        frmC.Left = esq + imgFec(Index).Parent.Left + 30
        frmC.Top = dalt + imgFec(Index).Parent.Top + imgFec(Index).Height + menu - 40
'    Else 'esquerra i dalt
'        frmC.Left = esq + btnFec(Index).Parent.Left - frmC.Width + btnFec(Index).Width + 40
'        frmC.Top = dalt + btnFec(Index).Parent.Top - frmC.Height + menu - 25
'    End If
    ' ***************

    imgFec(11).Tag = Index '<===
    ' *** repasar si el camp es txtAux o Text1 ***
    If Text1(Index).Text <> "" Then frmC.NovaData = Text1(Index).Text
    ' ********************************************

    frmC.Show vbModal
    Set frmC = Nothing
    ' *** repasar si el camp es txtAux o Text1 ***
    PonerFoco Text1(CByte(imgFec(11).Tag)) '<===
    ' ********************************************
End Sub

Private Sub frmC_Selec(vFecha As Date)
    ' *** repasar si el camp es txtAux o Text1 ***
    Text1(CByte(imgFec(11).Tag)).Text = Format(vFecha, "dd/mm/yyyy") '<===
    ' ********************************************
End Sub
' *****************************************************

Private Sub mnBuscar_Click()
    BotonBuscar
    Combo1(0).ListIndex = -1 'quan busque, per defecte no seleccione cap tipo de client
End Sub

Private Sub mnEliminar_Click()
    BotonEliminar
End Sub

Private Sub mnImprimir_Click()
    AbrirListado (11)
End Sub

Private Sub mnModificar_Click()
    ' *** repasar el nom de l'adodc, l'index del Field i el camp que te la PK ***
    'El registre de codi 0 no es pot Modificar ni Eliminar
    If EsCodigoCero(CStr(Data1.Recordset.Fields(0).Value), FormatoCampo(Text1(0))) Then Exit Sub
    ' ***************************************************************************
    
    If BLOQUEADesdeFormulario2(Me, Data1, 1) Then BotonModificar
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
        Case 12 'Imprimir
            mnImprimir_Click
'            printNou
        Case 13    'Eixir
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

    CadB = ObtenerBusqueda2(Me, BuscaChekc, 1)
    
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
    cad = cad & ParaGrid(Text1(2), 25, "EAN")
    If cad <> "" Then
        Screen.MousePointer = vbHourglass
        Set frmB = New frmBuscaGrid
        frmB.vCampos = cad
        frmB.vTabla = NombreTabla
        frmB.vSQL = CadB
        HaDevueltoDatos = False
        frmB.vDevuelve = "0|1|2|" '*** els camps que volen que torne ***
        frmB.vTitulo = "Articulos" ' ***** repasa açò: títol de BuscaGrid *****
        frmB.vSelElem = 1

        frmB.Show vbModal
        Set frmB = Nothing
        'Si ha posat valors i tenim que es formulari de búsqueda llavors
        'tindrem que tancar el form llançant l'event
        If HaDevueltoDatos Then
            If (Not Data1.Recordset.EOF) And DatosADevolverBusqueda <> "" Then _
                CmdRegresar_Click
        Else   'de ha retornat datos, es a decir NO ha retornat datos
            PonerFoco Text1(kCampo)
        End If
    End If
End Sub

Private Sub MandaBusquedaArticulo(CadB As String)
    Dim cad As String
        
    'Cridem al form
    ' **************** arreglar-ho per a vore lo que es desije ****************
    ' NOTA: el total d'amples de ParaGrid, ha de sumar 100
    cad = ""
    cad = cad & ParaGrid(Text1(0), 15, "Cód.")
    cad = cad & ParaGrid(Text1(1), 60, "Nombre")
    cad = cad & ParaGrid(Text1(2), 25, "EAN")
    If cad <> "" Then
        Screen.MousePointer = vbHourglass
        Set frmB1 = New frmBuscaGrid
        frmB1.vCampos = cad
        frmB1.vTabla = NombreTabla
        frmB1.vSQL = CadB
        HaDevueltoDatos = False
        frmB1.vDevuelve = "0|1|2|" '*** els camps que volen que torne ***
        frmB1.vTitulo = "Articulos" ' ***** repasa açò: títol de BuscaGrid *****
        frmB1.vSelElem = 1

        frmB1.Show vbModal
        Set frmB1 = Nothing
        'Si ha posat valors i tenim que es formulari de búsqueda llavors
        'tindrem que tancar el form llançant l'event
        If HaDevueltoDatos Then
            If (Not Data1.Recordset.EOF) And DatosADevolverBusqueda <> "" Then _
                CmdRegresar_Click
        Else   'de ha retornat datos, es a decir NO ha retornat datos
            PonerFoco Text1(19)
        End If
    End If
End Sub

Private Sub CmdRegresar_Click()
Dim cad As String
Dim Aux As String
Dim i As Integer
Dim J As Integer

    If Data1.Recordset.EOF Then
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
            cad = cad & Text1(J).Text & "|"
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
        PonerModo 2
        'Data1.Recordset.MoveLast
        Data1.Recordset.MoveFirst
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
    Text1(0).Text = SugerirCodigoSiguienteStr("sartic", "codartic")
    FormateaCampo Text1(0)
       
    PonerFoco Text1(0) '*** 1r camp visible que siga PK ***
    PosicionarCombo Combo1(0), 0
    
    ' *** si n'hi han camps de descripció a la capçalera ***
    'PosarDescripcions
    Text1(18).Text = 0

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
    'El registre de codi 0 no es pot Modificar ni Eliminar
    If EsCodigoCero(CStr(Data1.Recordset.Fields(0).Value), FormatoCampo(Text1(0))) Then Exit Sub
    ' ***************************************************************************

    ' *************** canviar la pregunta ****************
    cad = "¿Seguro que desea eliminar el Articulo?"
    cad = cad & vbCrLf & "Código: " & Format(Data1.Recordset.Fields(0), FormatoCampo(Text1(0)))
    cad = cad & vbCrLf & "Nombre: " & Data1.Recordset.Fields(1)
    
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
    If Err.Number > 0 Then MuestraError Err.Number, "Eliminar Articulo", Err.Description
End Sub

Private Sub PonerCampos()
Dim i As Integer
Dim codPobla As String, desPobla As String
Dim CPostal As String, desProvi As String, desPais As String

    If Data1.Recordset.EOF Then Exit Sub
    PonerCamposForma2 Me, Data1, 1 'opcio=1: posa el format o els camps de la capçalera
    
    ' *** si n'hi han llínies en datagrids ***
    'For i = 0 To DataGridAux.Count - 1
    For i = 0 To 1
        CargaGrid i, True
        If Not AdoAux(i).Recordset.EOF Then _
            PonerCamposForma2 Me, AdoAux(i), 2, "FrameAux" & i
    Next i
    
    ' ************* configurar els camps de les descripcions de la capçalera *************
    text2(3).Text = PonerNombreDeCod(Text1(3), "sfamia", "nomfamia", "codfamia", "N")
    text2(4).Text = PonerNombreCuenta(Text1(4), Modo)
    text2(17).Text = PonerNombreCuenta(Text1(17), Modo)
    text2(20).Text = PonerNombreCuenta(Text1(20), Modo) ' cuenta contable de compras
    'VRS:4.0.1(1)
    text2(0).Text = PonerNombreDeCod(Text1(5), "tiposiva", "nombriva", "codigiva", "N", cConta)
    text2(19).Text = PonerNombreDeCod(Text1(19), "sartic", "nomartic", "codartic", "N")

    '-- Esto permanece para saber donde estamos
    lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
    
    
    Me.chkDomicilio(0).visible = (Combo1(0).ListIndex = 3)
    Me.chkDomicilio(0).Enabled = (Combo1(0).ListIndex = 3) And Modo <> 0 And Modo <> 2
    If Not chkDomicilio(0).visible Then
        chkDomicilio(0).Value = 0
    End If
    
    
'    PonerModoOpcionesMenu (Modo)
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
Dim Mens As String


    On Error GoTo EDatosOK

    DatosOk = False
    b = CompForm2(Me, 1)
    If Not b Then Exit Function
    
    ' *** canviar els arguments de la funcio, el mensage i repasar si n'hi ha codEmpre ***
    If (Modo = 3) Then 'insertar
        'comprobar si existe ya el cod. del campo clave primaria
        If ExisteCP(Text1(0)) Then b = False
    End If
    If Not b Then Exit Function
    ' ************************************************************************************
    
    ' comprobamos que si hay contabilidad las cuentas contables existan
    If Modo = 3 Or Modo = 4 Then
        If vParamAplic.NumeroConta <> 0 Then
            text2(4).Text = PonerNombreCuenta(Text1(4), Modo)
            text2(17).Text = PonerNombreCuenta(Text1(17), Modo)
            text2(20).Text = PonerNombreCuenta(Text1(20), Modo)
            
            '[Monica]03/11/2014: en el caso de que cta cliente y cta socio sean las mismas damos un aviso pero dejamos continuar
            If Text1(4).Text = Text1(17).Text Then
                Mens = "Las cuentas de socio y de cliente coinciden. ¿ Desea continuar ?"
                If MsgBox(Mens, vbQuestion + vbYesNo + vbDefaultButton1) = vbNo Then b = False
            End If
                 
        End If
    End If
    
    DatosOk = b
    
EDatosOK:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Function

Private Sub PosicionarData()
Dim cad As String, Indicador As String

    ' *** canviar-ho per tota la PK de la capçalera, no llevar els () ***
    cad = "(codartic=" & Text1(0).Text & ")"
    
    ' *** gastar SituarData o SituarDataMULTI depenent de si la PK es simple o composta ***
    'If SituarDataMULTI(Data1, cad, Indicador) Then
    If SituarData(Data1, cad, Indicador) Then
        If ModoLineas <> 1 Then PonerModo 2
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
    vWhere = " WHERE codartic=" & Data1.Recordset!codArtic
        
    ' ***** elimina les llínies ****
    Conn.Execute "DELETE FROM starif " & vWhere
        
    Conn.Execute "DELETE FROM sbonif " & vWhere
        
    'Eliminar la CAPÇALERA
    vWhere = " WHERE codartic=" & Data1.Recordset!codArtic
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

Private Sub Text1_LostFocus(Index As Integer)
Dim cadMen As String
Dim Nuevo As Boolean
Dim SQL As String

    If Not PerderFocoGnral(Text1(Index), Modo) Then Exit Sub
    
    If Modo = 1 Then Exit Sub
    
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub
    
    ' ***************** configurar els LostFocus dels camps de la capçalera *****************
    Select Case Index
        Case 0, 5, 14, 15 'CODIGO
            PonerFormatoEntero Text1(Index)
            
            If Index = 5 And Text1(5).Text <> "" Then
                text2(0).Text = DevuelveDesdeBDNew(cConta, "tiposiva", "nombriva", "codigiva", Text1(5).Text, "N")
                If text2(0).Text = "" Then
                    MsgBox "Código de Iva no existe. Reintroduzca.", vbExclamation
                    PonerFoco Text1(Index)
                End If
            End If

        Case 6, 8, 13  'STOCKS Y PRECIOS
'### [Monica] 25/09/2006
'            cadMen = TransformaPuntosComas(Text1(Index).Text)
'            Text1(Index).Text = Format(cadMen, "#,###,##0.000")
            PonerFormatoDecimal Text1(Index), 5
            
        Case 9, 10 ' precios de compras
            PonerFormatoDecimal Text1(Index), 7
            
        Case 7 'BONIFICACIONES
            cadMen = TransformaPuntosComas(Text1(Index).Text)
            Text1(Index).Text = Format(cadMen, "#.0000")
        
        Case 1 'NOMBRE
            Text1(Index).Text = UCase(Text1(Index).Text)
        
        Case 3 'FAMILIA
            If PonerFormatoEntero(Text1(Index)) Then
                text2(Index).Text = PonerNombreDeCod(Text1(Index), "sfamia", "nomfamia")
                If text2(Index).Text = "" Then
                    cadMen = "No existe la Familia: " & Text1(Index).Text & vbCrLf
                    cadMen = cadMen & "¿Desea crearla?" & vbCrLf
                    If MsgBox(cadMen, vbQuestion + vbYesNo) = vbYes Then
                        Set frmFam = New frmManFamia
                        frmFam.DatosADevolverBusqueda = "0|1|"
                        frmFam.NuevoCodigo = Text1(Index).Text
                        Text1(Index).Text = ""
                        TerminaBloquear
                        frmFam.Show vbModal
                        Set frmFam = Nothing
                        If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Data1, 1
                    Else
                        Text1(Index).Text = ""
                    End If
                    PonerFoco Text1(Index)
                End If
                ' ### [Monica] 04/09/2006
                BloquearDatosCombustible Modo
            Else
                text2(Index).Text = ""
            End If
        
        Case 11, 12 'Fechas
            PonerFormatoFecha Text1(Index)
        
        Case 4, 17 'cuenta contable
            If Text1(Index).Text = "" Then Exit Sub
            text2(Index).Text = PonerNombreCuenta(Text1(Index), Modo)

        Case 20 ' cta contable de compra
            If Text1(Index).Text = "" Then Exit Sub
            text2(Index).Text = PonerNombreCuenta(Text1(Index), Modo)

'### [Monica] 25/09/2006
        Case 16
            If Text1(Index).Text <> "" Then PonerFormatoDecimal Text1(Index), 6
        
        Case 18
            If Text1(Index).Text <> "" Then PonerFormatoDecimal Text1(Index), 4
        
        Case 19
            If Text1(Index).Text <> "" Then
                PonerFormatoEntero Text1(Index)
                text2(Index).Text = DevuelveDesdeBDNew(cPTours, "sartic", "nomartic", "codartic", Text1(19).Text, "N")
                SQL = DevuelveDesdeBDNew(cPTours, "sartic", "codfamia", "codartic", Text1(19).Text, "N")
                If SQL = "" Then
                    MsgBox "Código de artículo no existente. Revise.", vbExclamation
                Else
                    If CInt(SQL) <> CInt(vParamAplic.FamDto) And vParamAplic.FamDto <> 0 Then
                        MsgBox "El Artículo de descuento debe pertenecer a la familia de descuento.", vbExclamation
                        Text1(19).Text = ""
                        text2(19).Text = ""
                        PonerFoco Text1(19)
                    End If
                End If
            End If
            
        Case 21 'peso articulo
            PonerFormatoDecimal Text1(Index), 4
        
        Case 22 'precio sigaus
            PonerFormatoDecimal Text1(Index), 6
            
    End Select
        ' ***************************************************************************
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = teclaBuscar Then
        If Modo = 1 Or Modo = 3 Or Modo = 4 Then
            Select Case Index
                Case 3: KEYBusqueda KeyAscii, 0 'familia
                Case 4: KEYBusqueda KeyAscii, 1 'cuenta de socio
                Case 17: KEYBusqueda KeyAscii, 2 'cuenta cliente
                Case 5: KEYBusqueda KeyAscii, 3 'tipo de iva
                Case 12: KEYFecha KeyAscii, 12 'fecha de inventario
                Case 11: KEYFecha KeyAscii, 11 'fecha de ultimo movimiento
                Case 20: KEYBusqueda KeyAscii, 5 'cuenta contable de compra
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

Private Sub KEYBusquedaLin(KeyAscii As Integer, indice As Integer)
    KeyAscii = 0
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
        Case 0 'tarifas
            SQL = "¿Seguro que desea eliminar la Tarifa?"
            SQL = SQL & vbCrLf & "Tarifa: " & AdoAux(Index).Recordset!codtarif
            If MsgBox(SQL, vbQuestion + vbYesNo) = vbYes Then
                eliminar = True
                SQL = "DELETE FROM starif"
                SQL = SQL & vWhere & " AND codtarif= " & AdoAux(Index).Recordset!codtarif
            End If
            
        Case 1 'bonificaciones
            SQL = "¿Seguro que desea eliminar la Bonificacion?"
            SQL = SQL & vbCrLf & "No.: " & AdoAux(Index).Recordset!NumLinea
            If MsgBox(SQL, vbQuestion + vbYesNo) = vbYes Then
                eliminar = True
                SQL = "DELETE FROM sbonif"
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
        Case 0: vTabla = "starif"
        Case 1: vTabla = "sbonif"
    End Select
    
    vWhere = ObtenerWhereCab(False)
    
    Select Case Index
        Case 0, 1 ' *** pose els index dels tabs de llínies que tenen datagrid ***
            ' *** canviar la clau primaria de les llínies,
            'pasar a "" si no volem que mos sugerixca res a l'afegir ***
            If Index = 1 Then
                NumF = SugerirCodigoSiguienteStr(vTabla, "numlinea", vWhere)
            End If

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
                Case 0 'tarifas
                    txtAux(0).Text = Text1(0).Text 'codartic
                    txtAux(1).Text = ""
                    txtAux(2).Text = ""
                    
                    PonerFoco txtAux(1)
                    
                Case 1 'bonificaciones
                    txtAux(8).Text = Text1(0).Text 'codartic
                    txtAux(9).Text = NumF 'numlinea
                    For i = 10 To 12
                        txtAux(i).Text = ""
                    Next i
                    cmbAux(0).ListIndex = -1
                    
            End Select
    End Select
End Sub

Private Sub BotonModificarLinea(Index As Integer)
    Dim anc As Single
    Dim i As Integer
    Dim J As Integer
    
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
        Case 0, 1 ' *** pose els index de llínies que tenen datagrid (en o sense tab) ***
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
        Case 0 'TARIFAS
        
            For J = 0 To 2
                txtAux(J).Text = DataGridAux(Index).Columns(J).Text
            Next J
            
            For i = 0 To 0
                BloquearTxt txtAux(i), False
            Next i
            
        Case 1 'BONIFICACIONES
            For J = 8 To 9
                txtAux(J).Text = DataGridAux(Index).Columns(J - 8).Text
            Next J
            
            For J = 10 To 12
                txtAux(J).Text = DataGridAux(Index).Columns(J - 6).Text
            Next J
            
            PosicionarCombo cmbAux(0), AdoAux(Index).Recordset!tipsocio
            
            For i = 8 To 9
                BloquearTxt txtAux(i), False
            Next i
            
    End Select
    
    LLamaLineas Index, ModoLineas, anc
   
    ' *** foco al 1r camp visible de les llinies en grids que no siga PK (en o sense tab) ***
    Select Case Index
        Case 0 'tarifas
            PonerFoco txtAux(1)
        Case 1 'bonificaciones
            PonerFocoCmb cmbAux(0)
    End Select
    ' ***************************************************************************************
End Sub

Private Sub LLamaLineas(Index As Integer, xModo As Byte, Optional alto As Single)
Dim jj As Integer
Dim b As Boolean

    ' *** si n'hi han tabs sense datagrid posar el If ***
    DeseleccionaGrid DataGridAux(Index)
       
    b = (xModo = 1 Or xModo = 2) 'Insertar o Modificar Llínies
    Select Case Index
        Case 0 'tarifas
             For jj = 1 To 2
                txtAux(jj).visible = b
                txtAux(jj).Top = alto
            Next jj
            
            
        Case 1 'matriculas
            For jj = 10 To 12
                txtAux(jj).visible = b
                txtAux(jj).Top = alto
            Next jj
            
            cmbAux(0).visible = b
            cmbAux(0).Top = alto - 15
    End Select
End Sub

' ********* si n'hi han combos a la capçalera ************
Private Sub CargaCombo(Index As Integer)
Dim Ini As Integer
Dim Fin As Integer
Dim i As Integer

    ' *** neteje els combos, els pose valor i seleccione el valor per defecte ***
    cmbAux(0).Clear
    Combo1(0).Clear
    
    cmbAux(0).AddItem "Comercio"
    cmbAux(0).ItemData(cmbAux(0).NewIndex) = 2
    cmbAux(0).AddItem "Particular"
    cmbAux(0).ItemData(cmbAux(0).NewIndex) = 0
    cmbAux(0).AddItem "Profesional"
    cmbAux(0).ItemData(cmbAux(0).NewIndex) = 1
    
    Combo1(0).AddItem "Nada"
    Combo1(0).ItemData(Combo1(0).NewIndex) = 0
    Combo1(0).AddItem "Gasolina"
    Combo1(0).ItemData(Combo1(0).NewIndex) = 1
    Combo1(0).AddItem "Gasoleo A"
    Combo1(0).ItemData(Combo1(0).NewIndex) = 2
    Combo1(0).AddItem "Gasoleo B"
    Combo1(0).ItemData(Combo1(0).NewIndex) = 3
    Combo1(0).AddItem "Gasoleo C"
    Combo1(0).ItemData(Combo1(0).NewIndex) = 4
    
End Sub

Private Sub txtAux_LostFocus(Index As Integer)
Dim cadMen As String
Dim Nuevo As Boolean
    
    If Not PerderFocoGnral(txtAux(Index), Modo) Then Exit Sub
    
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub
    

    ' ******* configurar el LostFocus dels camps de llínies (dins i fora grid) ********
    Select Case Index
            
        Case 2 'PRECIO
            '### [Monica] 25/09/2006 comento las lineas siguientes
'            cadMen = TransformaPuntosComas(txtAux(Index).Text)
'            txtAux(Index).Text = Format(cadMen, "#,###,##0.000")
            If txtAux(Index).Text <> "" Then PonerFormatoDecimal txtAux(Index), 5
            PonerFocoBtn Me.cmdAceptar
        
        Case 10, 11 'DESDE - HASTA
            cadMen = TransformaPuntosComas(txtAux(Index).Text)
            txtAux(Index).Text = Format(cadMen, "#,###,###,##0")
        
        Case 12 'BONIFICACION
            '### [Monica] 25/09/2006 comento las lineas siguientes
'            cadMen = TransformaPuntosComas(txtAux(Index).Text)
'            txtAux(Index).Text = Format(cadMen, "##0.000")
            If txtAux(Index).Text <> "" Then PonerFormatoDecimal txtAux(Index), 2
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
Dim RS As ADODB.Recordset
Dim SQL As String
Dim b As Boolean
Dim Cant As Integer
Dim Mens As String
Dim vFact As Byte, vDocum As Byte

    On Error GoTo EDatosOKLlin

    Mens = ""
    DatosOkLlin = False
        
    b = CompForm2(Me, 2, nomFrame) 'Comprovar formato datos ok
    If Not b Then Exit Function
    
'    ' *** si cal fer atres comprovacions a les llínies (en o sense tab) ***
'    Select Case NumTabMto
'        Case 0  'CUENTAS BANCARIAS
'            SQL = "SELECT COUNT(ctaprpal) FROM cltebanc "
'            SQL = SQL & ObtenerWhereCab(True) & " AND ctaprpal=1"
'            If ModoLineas = 2 Then SQL = SQL & " AND numlinea<> " & AdoAux(NumTabMto).Recordset!numlinea
'            Set RS = New ADODB.Recordset
'            RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
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
'                If ModoLineas = 2 Then SQL = SQL & " AND numlinea<> " & AdoAux(NumTabMto).Recordset!numlinea
'                SQL = SQL & " AND codnacio=" & cmbAux(0).ItemData(cmbAux(0).ListIndex)
'                SQL = SQL & " AND codbanco=" & txtAux(3).Text & " AND codsucur=" & txtAux(4).Text
'                SQL = SQL & " AND digcontr='" & txtAux(5).Text & "' AND ctabanco='" & txtAux(6).Text & "'"
'                Set RS = New ADODB.Recordset
'                RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
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
'                    If ModoLineas = 2 Then SQL = SQL & " AND numlinea<> " & AdoAux(NumTabMto).Recordset!numlinea
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
'                    RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
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
'                RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
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
'                RS.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
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
        Case 0 'familias
            Set frmFam = New frmManFamia
            frmFam.DatosADevolverBusqueda = "0|1|"
            frmFam.CodigoActual = Text1(3).Text
            frmFam.Show vbModal
            Set frmFam = Nothing
            PonerFoco Text1(3)
            
        Case 1, 2 'Cuentas Contables (de contabilidad)
            If vParamAplic.NumeroConta = 0 Then Exit Sub
            If Index = 1 Then
                indice = 4
            Else
                indice = 17
            End If
            Set frmCtas = New frmCtasConta
            frmCtas.NumDigit = 0
            frmCtas.DatosADevolverBusqueda = "0|1|"
            frmCtas.CodigoActual = Text1(indice).Text
            frmCtas.Show vbModal
            Set frmCtas = Nothing
            PonerFoco Text1(indice)
        Case 5 ' Cuenta contable de compras
            If vParamAplic.NumeroConta = 0 Then Exit Sub
            indice = 20
            Set frmCtas = New frmCtasConta
            frmCtas.NumDigit = 0
            frmCtas.DatosADevolverBusqueda = "0|1|"
            frmCtas.CodigoActual = Text1(indice).Text
            frmCtas.Show vbModal
            Set frmCtas = Nothing
            PonerFoco Text1(indice)
        
        
        Case 3 'VRS:4.0.1
            Set frmTipIva = New frmTipIVAConta
            frmTipIva.DatosADevolverBusqueda = "0|1|2|"
            frmTipIva.CodigoActual = Text1(5).Text
            frmTipIva.Show vbModal
            Set frmTipIva = Nothing
        
        Case 4 'Articulo de descuento
            MandaBusquedaArticulo ""
    End Select
    
    If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Data1, 1
End Sub

Private Sub frmCtas_DatoSeleccionado(CadenaSeleccion As String)
'Cuentas contables de la Contabilidad
    Text1(indice).Text = RecuperaValor(CadenaSeleccion, 1) 'codmacta
    text2(indice).Text = RecuperaValor(CadenaSeleccion, 2) 'des macta
End Sub

Private Sub frmFam_DatoSeleccionado(CadenaSeleccion As String)
'Mantenimiento de Colectivos
    Text1(3).Text = RecuperaValor(CadenaSeleccion, 1) 'codfamia
    FormateaCampo Text1(3)
    text2(3).Text = RecuperaValor(CadenaSeleccion, 2) 'nomfamia
End Sub

Private Sub DataGridAux_RowColChange(Index As Integer, LastRow As Variant, ByVal LastCol As Integer)
Dim i As Byte

    If ModoLineas <> 1 Then
        Select Case Index
            Case 0 'tarifas
                If DataGridAux(Index).Columns.Count > 2 Then
'                    txtAux(11).Text = DataGridAux(Index).Columns("direccio").Text
'                    txtAux(12).Text = DataGridAux(Index).Columns("observac").Text
                End If
                
            Case 1 'bonificaiones
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
            Case 0 'cuentas bancarias
'                txtAux(11).Text = ""
'                txtAux(12).Text = ""
            Case 1 'departamentos
                For i = 21 To 24
'                   txtAux(i).Text = ""
                Next i
'               txtAux2(22).Text = ""
            Case 2 'Tarjetas
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
        Case 0 'tarifas
            'si es visible|control|tipo campo|nombre campo|ancho control|
            tots = "N||||0|;" 'codartic
            tots = tots & "S|txtAux(1)|T|Tarifa|650|;"
            tots = tots & "S|txtAux(2)|T|P.V.P.|1500|;"
            
            arregla tots, DataGridAux(Index), Me
        
'            DataGridAux(0).Columns(5).Alignment = dbgRight
'            DataGridAux(0).Columns(6).Alignment = dbgRight
'            DataGridAux(0).Columns(7).Alignment = dbgRight
'            DataGridAux(0).Columns(8).Alignment = dbgRight
        
            b = (Modo = 4) And ((ModoLineas = 1) Or (ModoLineas = 2))
            
        Case 1 'Bonificaciones
            'si es visible|control|tipo campo|nombre campo|ancho control|
            tots = "N||||0|;N||||0|;" 'codartic,numlinea
            tots = tots & "N||||0|;S|cmbAux(0)|C|Tipo|1200|;"
            tots = tots & "S|txtAux(10)|T|Desde|1200|;S|txtAux(11)|T|Hasta|1200|;"
            tots = tots & "S|txtAux(12)|T|Bonificacion|1200|;"
            
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
        Case 0: nomFrame = "FrameAux0" 'tarifas
        Case 1: nomFrame = "FrameAux1" 'bonificaciones
    End Select
    
    If DatosOkLlin(nomFrame) Then
        TerminaBloquear
        If InsertarDesdeForm2(Me, 2, nomFrame) Then
            b = BLOQUEADesdeFormulario2(Me, Data1, 1)
            Select Case NumTabMto
                Case 0, 1 ' *** els index de les llinies en grid (en o sense tab) ***
                     CargaGrid NumTabMto, True
                    If b Then BotonAnyadirLinea NumTabMto
            End Select
           
            SituarTab (NumTabMto + 1)
        End If
    End If
End Sub

Private Sub ModificarLinea()
'Modifica registre en les taules de Llínies
Dim nomFrame As String
Dim V As Integer
    
    On Error Resume Next

    ' *** posa els noms del frames, tant si son de grid com si no ***
    Select Case NumTabMto
        Case 0: nomFrame = "FrameAux0" 'tarifas
        Case 1: nomFrame = "FrameAux1" 'bonificaciones
    End Select
    
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
        End If
    End If
End Sub

Private Function ObtenerWhereCab(conW As Boolean) As String
Dim vWhere As String
    
    vWhere = ""
    If conW Then vWhere = " WHERE "
    ' *** canviar-ho per la clau primaria de la capçalera ***
    vWhere = vWhere & " codartic=" & Val(Text1(0).Text)
    
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

Private Sub printNou()
    With frmImprimir2
        .cadTabla2 = "articulos"
        .Informe2 = "rClientes.rpt"
        If CadB <> "" Then
            .cadRegSelec = Replace(SQL2SF(CadB), "clientes", "clientes_1")
        Else
            .cadRegSelec = ""
        End If
        .cadRegActua = Replace(POS2SF(Data1, Me), "clientes", "clientes_1")
        .cadTodosReg = ""
        '.OtrosParametros2 = "pEmpresa='" & vEmpresa.NomEmpre & "'|pOrden={clientes.ape_raso}|"
        .OtrosParametros2 = "pEmpresa='" & vEmpresa.nomEmpre & "'|"
        .NumeroParametros2 = 1
        .MostrarTree2 = False
        .InfConta2 = False
        .ConSubInforme2 = False
        
        .Show vbModal
    End With
End Sub

' ### [DavidV] 26/04/2006: Activar/desactivar la rueda del ratón.
Private Sub DataGridAux_GotFocus(Index As Integer)
  WheelHook DataGridAux(Index)
End Sub
Private Sub DataGridAux_LostFocus(Index As Integer)
  WheelUnHook
End Sub


' ### [Monica] 04/09/2006
Private Sub BloquearDatosCombustible(Modo)
'Activem el frame de dades de combustible
Dim b As Boolean, bAux As Boolean
Dim i As Byte
    
    'Barra de CAPÇALERA
    '------------------------------------------
    'b = (Modo = 2 Or Modo = 0 Or Modo = 1)
    b = (Modo < 4) ' si es distinto de modificar y no estamos en lineas está activo
    Frame4.Enabled = b
    
    b = (Modo = 4) Or (Modo = 3 And Text1(3).Text <> "")
    If b Then
        i = DevuelveDesdeBDNew(cPTours, "sfamia", "tipfamia", "codfamia", Text1(3).Text, "N")
        Frame4.Enabled = (i = 1)
    End If
    
End Sub

