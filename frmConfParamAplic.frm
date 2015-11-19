VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmConfParamAplic 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Parámetros de la Aplicación"
   ClientHeight    =   7125
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   8685
   Icon            =   "frmConfParamAplic.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7125
   ScaleWidth      =   8685
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   5865
      Left            =   180
      TabIndex        =   39
      Top             =   630
      Width           =   8085
      _ExtentX        =   14261
      _ExtentY        =   10345
      _Version        =   393216
      Tabs            =   5
      Tab             =   4
      TabsPerRow      =   5
      TabHeight       =   520
      ForeColor       =   8388608
      TabCaption(0)   =   "Contabilidad"
      TabPicture(0)   =   "frmConfParamAplic.frx":000C
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame3"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Otros Datos"
      TabPicture(1)   =   "frmConfParamAplic.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame5"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "FrameCooperativa"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "FrameHidrocarburos"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "Internet"
      TabPicture(2)   =   "frmConfParamAplic.frx":0044
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame4"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Frame7"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).ControlCount=   2
      TabCaption(3)   =   "Compras/Internas"
      TabPicture(3)   =   "frmConfParamAplic.frx":0060
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "imgAyuda(0)"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "Label1(8)"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).Control(2)=   "Label1(9)"
      Tab(3).Control(2).Enabled=   0   'False
      Tab(3).Control(3)=   "Text1(34)"
      Tab(3).Control(3).Enabled=   0   'False
      Tab(3).Control(4)=   "Frame6"
      Tab(3).Control(4).Enabled=   0   'False
      Tab(3).Control(5)=   "chkctrstock"
      Tab(3).Control(5).Enabled=   0   'False
      Tab(3).Control(6)=   "Frame8"
      Tab(3).Control(6).Enabled=   0   'False
      Tab(3).Control(7)=   "Combo2"
      Tab(3).Control(7).Enabled=   0   'False
      Tab(3).ControlCount=   8
      TabCaption(4)   =   "Facturación"
      TabPicture(4)   =   "frmConfParamAplic.frx":007C
      Tab(4).ControlEnabled=   -1  'True
      Tab(4).Control(0)=   "Frame9"
      Tab(4).Control(0).Enabled=   0   'False
      Tab(4).Control(1)=   "Frame10"
      Tab(4).Control(1).Enabled=   0   'False
      Tab(4).ControlCount=   2
      Begin VB.Frame Frame10 
         Caption         =   "Códigos de Iva Antiguo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1785
         Left            =   420
         TabIndex        =   120
         Top             =   3150
         Width           =   7185
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   43
            Left            =   1830
            MaxLength       =   3
            TabIndex        =   124
            Tag             =   "Iva Reducido Antiguo|N|N|||sparam|codiva2old|||"
            Top             =   1020
            Width           =   615
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Enabled         =   0   'False
            Height          =   285
            Index           =   42
            Left            =   2520
            TabIndex        =   123
            Top             =   510
            Width           =   3345
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   42
            Left            =   1830
            MaxLength       =   3
            TabIndex        =   122
            Tag             =   "Iva General Antiguo|N|N|||sparam|codiva1old|||"
            Top             =   510
            Width           =   615
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Enabled         =   0   'False
            Height          =   285
            Index           =   43
            Left            =   2520
            TabIndex        =   121
            Top             =   1020
            Width           =   3345
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   18
            Left            =   1500
            ToolTipText     =   "Buscar Iva"
            Top             =   540
            Width           =   255
         End
         Begin VB.Label Label1 
            Caption         =   "General"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   34
            Left            =   270
            TabIndex        =   126
            Top             =   540
            Width           =   1095
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   19
            Left            =   1500
            ToolTipText     =   "Buscar Iva"
            Top             =   1050
            Width           =   255
         End
         Begin VB.Label Label1 
            Caption         =   "Reducido"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   33
            Left            =   270
            TabIndex        =   125
            Top             =   1050
            Width           =   1095
         End
      End
      Begin VB.Frame Frame9 
         Caption         =   "Códigos de Iva"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2235
         Left            =   420
         TabIndex        =   110
         Top             =   720
         Width           =   7185
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Enabled         =   0   'False
            Height          =   285
            Index           =   41
            Left            =   2520
            TabIndex        =   118
            Top             =   1500
            Width           =   3345
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   41
            Left            =   1830
            MaxLength       =   3
            TabIndex        =   117
            Tag             =   "Iva Super-Reducidol|N|N|||sparam|codiva3|||"
            Top             =   1500
            Width           =   615
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Enabled         =   0   'False
            Height          =   285
            Index           =   40
            Left            =   2520
            TabIndex        =   115
            Top             =   1020
            Width           =   3345
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   40
            Left            =   1830
            MaxLength       =   3
            TabIndex        =   114
            Tag             =   "Iva Reducido|N|N|||sparam|codiva2|||"
            Top             =   1020
            Width           =   615
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Enabled         =   0   'False
            Height          =   285
            Index           =   39
            Left            =   2520
            TabIndex        =   112
            Top             =   510
            Width           =   3345
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   39
            Left            =   1830
            MaxLength       =   3
            TabIndex        =   111
            Tag             =   "Iva General|N|N|||sparam|codiva1|||"
            Top             =   510
            Width           =   615
         End
         Begin VB.Label Label1 
            Caption         =   "Super-Reducido"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   31
            Left            =   270
            TabIndex        =   119
            Top             =   1530
            Width           =   1185
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   17
            Left            =   1500
            ToolTipText     =   "Buscar Iva"
            Top             =   1530
            Width           =   255
         End
         Begin VB.Label Label1 
            Caption         =   "Reducido"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   30
            Left            =   270
            TabIndex        =   116
            Top             =   1050
            Width           =   1095
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   16
            Left            =   1500
            ToolTipText     =   "Buscar Iva"
            Top             =   1050
            Width           =   255
         End
         Begin VB.Label Label1 
            Caption         =   "General"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   16
            Left            =   270
            TabIndex        =   113
            Top             =   540
            Width           =   1095
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   15
            Left            =   1500
            ToolTipText     =   "Buscar Iva"
            Top             =   540
            Width           =   255
         End
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   -69900
         Style           =   2  'Dropdown List
         TabIndex        =   95
         Tag             =   "Precio Artículo|N|N|||sparam|tipoprecio||N|"
         Top             =   1560
         Width           =   2340
      End
      Begin VB.Frame Frame7 
         Height          =   2565
         Left            =   -74880
         TabIndex        =   67
         Top             =   450
         Width           =   7875
         Begin VB.CheckBox chkOutlook 
            Caption         =   "Enviar desde Outlook"
            Height          =   375
            Left            =   5310
            TabIndex        =   104
            Tag             =   "Outlook|N|N|||sparam|EnvioDesdeOutlook|||"
            Top             =   2010
            Width           =   2175
         End
         Begin VB.TextBox Text1 
            Height          =   285
            IMEMode         =   3  'DISABLE
            Index           =   35
            Left            =   2730
            MaxLength       =   30
            TabIndex        =   31
            Tag             =   "LanzaMailOutlook|T|S|||sparam|arigesmail|||"
            Text            =   "3"
            Top             =   2040
            Width           =   1620
         End
         Begin VB.TextBox Text1 
            Height          =   285
            IMEMode         =   3  'DISABLE
            Index           =   19
            Left            =   5250
            MaxLength       =   15
            PasswordChar    =   "*"
            TabIndex        =   30
            Tag             =   "Password SMTP|T|S|||sparam|smtpPass|||"
            Text            =   "3"
            Top             =   1440
            Width           =   2220
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   18
            Left            =   1260
            MaxLength       =   50
            TabIndex        =   29
            Tag             =   "Usuario SMTP|T|S|||sparam|smtpUser|||"
            Text            =   "3"
            Top             =   1440
            Width           =   3090
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   17
            Left            =   1260
            MaxLength       =   50
            TabIndex        =   28
            Tag             =   "Servidor SMTP|T|S|||sparam|smtpHost|||"
            Text            =   "3"
            Top             =   900
            Width           =   6210
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   16
            Left            =   1260
            MaxLength       =   50
            TabIndex        =   27
            Tag             =   "Direccion e-mail|T|S|||sparam|diremail|||"
            Text            =   "3"
            Top             =   420
            Width           =   6210
         End
         Begin VB.Label Label1 
            Caption         =   "Lanza pantalla mail outlook"
            Height          =   195
            Index           =   60
            Left            =   120
            TabIndex        =   103
            Top             =   2070
            Width           =   2040
         End
         Begin VB.Label Label8 
            Caption         =   "Envio E-Mail"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   0
            Left            =   360
            TabIndex        =   72
            Top             =   0
            Width           =   1320
         End
         Begin VB.Label Label1 
            Caption         =   "Password"
            Height          =   195
            Index           =   23
            Left            =   4440
            TabIndex        =   71
            Top             =   1500
            Width           =   840
         End
         Begin VB.Label Label1 
            Caption         =   "Usuario"
            Height          =   195
            Index           =   22
            Left            =   120
            TabIndex        =   70
            Top             =   1500
            Width           =   1380
         End
         Begin VB.Label Label1 
            Caption         =   "Servidor SMTP"
            Height          =   195
            Index           =   21
            Left            =   120
            TabIndex        =   69
            Top             =   960
            Width           =   1380
         End
         Begin VB.Label Label1 
            Caption         =   "E-Mail"
            Height          =   195
            Index           =   20
            Left            =   120
            TabIndex        =   68
            Top             =   480
            Width           =   1380
         End
      End
      Begin VB.Frame Frame8 
         Caption         =   "Facturación Interna"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   -74640
         TabIndex        =   100
         Top             =   2460
         Width           =   7155
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   33
            Left            =   2475
            MaxLength       =   3
            TabIndex        =   96
            Tag             =   "Iva Exento|N|S|||sparam|tipoivaexento|||"
            Top             =   360
            Width           =   615
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Enabled         =   0   'False
            Height          =   285
            Index           =   33
            Left            =   3150
            TabIndex        =   101
            Top             =   360
            Width           =   3345
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   13
            Left            =   2100
            ToolTipText     =   "Buscar Iva"
            Top             =   360
            Width           =   255
         End
         Begin VB.Label Label1 
            Caption         =   "Código Iva Exento"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   12
            Left            =   240
            TabIndex        =   102
            Top             =   390
            Width           =   1845
         End
      End
      Begin VB.CheckBox chkctrstock 
         Caption         =   "Realiza control de Stock"
         Height          =   375
         Left            =   -74520
         TabIndex        =   94
         Tag             =   "Control de Stock|N|N|||sparam|ctrstock|||"
         Top             =   1500
         Width           =   2775
      End
      Begin VB.Frame Frame6 
         Caption         =   "Facturación Compras"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   -74640
         TabIndex        =   89
         Top             =   450
         Width           =   7185
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   320
            Index           =   31
            Left            =   6120
            MaxLength       =   2
            TabIndex        =   93
            Tag             =   "Mes a no girar|N|S|0|12|sparam|mesnogir|||"
            Text            =   "Text1"
            Top             =   360
            Width           =   615
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   320
            Index           =   30
            Left            =   2880
            MaxLength       =   2
            TabIndex        =   92
            Tag             =   "Dia 3 de pago compras|N|S|0|31|sparam|diapago3|||"
            Text            =   "Text1"
            Top             =   360
            Width           =   615
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   320
            Index           =   29
            Left            =   2160
            MaxLength       =   2
            TabIndex        =   91
            Tag             =   "Dia 2 de pago compras|N|S|0|31|sparam|diapago2|||"
            Text            =   "Text1"
            Top             =   360
            Width           =   615
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   320
            Index           =   28
            Left            =   1440
            MaxLength       =   2
            TabIndex        =   90
            Tag             =   "Dia 1 de pago compras|N|S|0|31|sparam|diapago1|||"
            Text            =   "Text1"
            Top             =   360
            Width           =   615
         End
         Begin VB.Label Label1 
            Caption         =   "Mes a no girar"
            Height          =   255
            Index           =   13
            Left            =   4920
            TabIndex        =   98
            Top             =   360
            Width           =   1095
         End
         Begin VB.Label Label1 
            Caption         =   "Días de pago"
            Height          =   255
            Index           =   11
            Left            =   360
            TabIndex        =   97
            Top             =   360
            Width           =   1095
         End
      End
      Begin VB.Frame FrameHidrocarburos 
         Caption         =   "Declaración de Hidrocarburos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   -74580
         TabIndex        =   75
         Top             =   3735
         Width           =   7095
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   21
            Left            =   2700
            MaxLength       =   4
            TabIndex        =   25
            Tag             =   "CEE|T|S|||sparam|cee|||"
            Text            =   "1234"
            Top             =   330
            Width           =   585
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   20
            Left            =   690
            MaxLength       =   8
            TabIndex        =   24
            Tag             =   "CIM|T|S|||sparam|cim|||"
            Text            =   "12345678"
            Top             =   330
            Width           =   945
         End
         Begin VB.Label Label4 
            Caption         =   "CEE"
            Height          =   255
            Left            =   2190
            TabIndex        =   77
            Top             =   360
            Width           =   555
         End
         Begin VB.Label Label3 
            Caption         =   "CIM"
            Height          =   255
            Left            =   180
            TabIndex        =   76
            Top             =   360
            Width           =   555
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Soporte"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1035
         Left            =   -74880
         TabIndex        =   73
         Top             =   3210
         Width           =   7845
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   14
            Left            =   1290
            MaxLength       =   100
            TabIndex        =   32
            Tag             =   "Web Soporte|T|S|||sparam|websoporte|||"
            Top             =   360
            Width           =   6135
         End
         Begin VB.Label Label2 
            Caption         =   "Web soporte"
            Height          =   255
            Left            =   180
            TabIndex        =   74
            Top             =   360
            Width           =   1575
         End
      End
      Begin VB.Frame FrameCooperativa 
         Caption         =   "Instalación "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   -74580
         TabIndex        =   66
         Top             =   4560
         Width           =   7095
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   15
            Left            =   360
            MaxLength       =   100
            TabIndex        =   26
            Tag             =   "Cooperativa|N|S|||sparam|cooperativa|00||"
            Top             =   360
            Width           =   705
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Otros"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3195
         Left            =   -74580
         TabIndex        =   61
         Top             =   480
         Width           =   7125
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   38
            Left            =   5370
            MaxLength       =   12
            TabIndex        =   17
            Tag             =   "Límite Facturas Efectivo|N|N|||sparam|limitefra|#,###,##0.00 ||"
            Top             =   615
            Width           =   1425
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   36
            Left            =   2220
            MaxLength       =   255
            TabIndex        =   23
            Tag             =   "Path FacturaE|T|S|||sparam|pathfacturae|||"
            Top             =   2760
            Width           =   4590
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   32
            Left            =   2220
            MaxLength       =   100
            TabIndex        =   22
            Tag             =   "Impresora Tarjetas|T|S|||sparam|impresoraticket|||"
            Top             =   2400
            Width           =   4590
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   27
            Left            =   2220
            MaxLength       =   3
            TabIndex        =   20
            Tag             =   "Familia Dto|N|S|0|999|sparam|famdto|000||"
            Top             =   1710
            Width           =   675
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Enabled         =   0   'False
            Height          =   285
            Index           =   27
            Left            =   2925
            TabIndex        =   87
            Top             =   1710
            Width           =   3885
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   26
            Left            =   2220
            MaxLength       =   50
            TabIndex        =   21
            Tag             =   "Impresora Tarjetas|T|S|||sparam|impresoratarjeta|||"
            Top             =   2055
            Width           =   4590
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Enabled         =   0   'False
            Height          =   285
            Index           =   25
            Left            =   2925
            TabIndex        =   85
            Top             =   1350
            Width           =   3885
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   25
            Left            =   2220
            MaxLength       =   6
            TabIndex        =   19
            Tag             =   "Cooperativa|N|N|0|99|sparam|coopdefecto|00||"
            Top             =   1350
            Width           =   675
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   13
            Left            =   2220
            MaxLength       =   6
            TabIndex        =   18
            Tag             =   "Código Artículo|N|S|0|999999|sparam|articdto|000000||"
            Top             =   990
            Width           =   1215
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Enabled         =   0   'False
            Height          =   285
            Index           =   7
            Left            =   3480
            TabIndex        =   62
            Top             =   990
            Width           =   3315
         End
         Begin VB.ComboBox Combo1 
            Height          =   315
            ItemData        =   "frmConfParamAplic.frx":0098
            Left            =   2220
            List            =   "frmConfParamAplic.frx":00A2
            Style           =   2  'Dropdown List
            TabIndex        =   16
            Tag             =   "Bonificacion|N|N|0|1|sparam|bonifact|||"
            Top             =   600
            Width           =   1095
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   12
            Left            =   2220
            MaxLength       =   20
            TabIndex        =   15
            Tag             =   "Texto Impuesto|T|S|||sparam|teximpue|||"
            Text            =   "3wwwwwwwwwwwwwwwwwww"
            Top             =   240
            Width           =   4590
         End
         Begin VB.Label Label7 
            Caption         =   "Límite Facturas Efectivo "
            Height          =   255
            Left            =   3420
            TabIndex        =   109
            Top             =   630
            Width           =   1755
         End
         Begin VB.Label Label1 
            Caption         =   "Path FacturaE"
            Height          =   195
            Index           =   10
            Left            =   210
            TabIndex        =   106
            Top             =   2805
            Width           =   2070
         End
         Begin VB.Label Label1 
            Caption         =   "Impresora Tickets"
            Height          =   195
            Index           =   7
            Left            =   210
            TabIndex        =   99
            Top             =   2445
            Width           =   2070
         End
         Begin VB.Label Label6 
            Caption         =   "Familia Descuento"
            Height          =   255
            Left            =   210
            TabIndex        =   88
            Top             =   1738
            Width           =   1575
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   12
            Left            =   1935
            ToolTipText     =   "Buscar Artículo"
            Top             =   1710
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Impresora Tarjetas"
            Height          =   195
            Index           =   6
            Left            =   210
            TabIndex        =   86
            Top             =   2120
            Width           =   2070
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   11
            Left            =   1935
            ToolTipText     =   "Buscar Artículo"
            Top             =   1350
            Width           =   240
         End
         Begin VB.Label Label5 
            Caption         =   "Colectivo Defecto"
            Height          =   255
            Left            =   210
            TabIndex        =   84
            Top             =   1356
            Width           =   1575
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   7
            Left            =   1935
            ToolTipText     =   "Buscar Artículo"
            Top             =   990
            Width           =   240
         End
         Begin VB.Label Label11 
            Caption         =   "Artículo descuento"
            Height          =   255
            Left            =   210
            TabIndex        =   65
            Top             =   974
            Width           =   1575
         End
         Begin VB.Label Label1 
            Caption         =   "Bonificación"
            Height          =   195
            Index           =   29
            Left            =   210
            TabIndex        =   64
            Top             =   652
            Width           =   2070
         End
         Begin VB.Label Label1 
            Caption         =   "Texto Impuesto"
            Height          =   195
            Index           =   28
            Left            =   210
            TabIndex        =   63
            Top             =   330
            Width           =   2070
         End
      End
      Begin VB.Frame Frame3 
         Height          =   3870
         Left            =   -74820
         TabIndex        =   46
         Top             =   1920
         Width           =   7665
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Enabled         =   0   'False
            Height          =   285
            Index           =   37
            Left            =   3645
            Locked          =   -1  'True
            TabIndex        =   107
            Top             =   3150
            Width           =   3435
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   37
            Left            =   2955
            MaxLength       =   10
            TabIndex        =   13
            Tag             =   "Concepto Haber|T|S|||sparam|concehaberresto|||"
            Top             =   3150
            Width           =   585
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   23
            Left            =   2955
            MaxLength       =   10
            TabIndex        =   12
            Tag             =   "Concepto Haber|T|S|||sparam|concehaber|||"
            Top             =   2835
            Width           =   585
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Enabled         =   0   'False
            Height          =   285
            Index           =   23
            Left            =   3645
            Locked          =   -1  'True
            TabIndex        =   80
            Top             =   2835
            Width           =   3435
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   22
            Left            =   2955
            MaxLength       =   10
            TabIndex        =   11
            Tag             =   "Concepto Debe|T|S|||sparam|concedebe|||"
            Top             =   2505
            Width           =   585
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Enabled         =   0   'False
            Height          =   285
            Index           =   22
            Left            =   3645
            Locked          =   -1  'True
            TabIndex        =   79
            Top             =   2505
            Width           =   3450
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   24
            Left            =   2955
            MaxLength       =   10
            TabIndex        =   14
            Tag             =   "Numero Diario|T|S|||sparam|numdiari|||"
            Top             =   3510
            Width           =   585
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Enabled         =   0   'False
            Height          =   285
            Index           =   24
            Left            =   3645
            Locked          =   -1  'True
            TabIndex        =   78
            Top             =   3510
            Width           =   3405
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Enabled         =   0   'False
            Height          =   285
            Index           =   0
            Left            =   4200
            TabIndex        =   53
            Top             =   240
            Width           =   2895
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   5
            Left            =   2940
            MaxLength       =   10
            TabIndex        =   4
            Tag             =   "Cta.Contable|T|S|||sparam|ctaconta|||"
            Top             =   240
            Width           =   1215
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Enabled         =   0   'False
            Height          =   285
            Index           =   1
            Left            =   4200
            TabIndex        =   52
            Top             =   570
            Width           =   2895
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   6
            Left            =   2940
            MaxLength       =   10
            TabIndex        =   5
            Tag             =   "Cta.Dif.negativas|T|S|||sparam|ctanegtat|||"
            Top             =   570
            Width           =   1215
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Enabled         =   0   'False
            Height          =   285
            Index           =   2
            Left            =   4200
            TabIndex        =   51
            Top             =   900
            Width           =   2895
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   7
            Left            =   2940
            MaxLength       =   10
            TabIndex        =   6
            Tag             =   "Cta.Dif.Positivas|T|S|||sparam|ctaposit|||"
            Top             =   900
            Width           =   1215
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Enabled         =   0   'False
            Height          =   285
            Index           =   3
            Left            =   4200
            TabIndex        =   50
            Top             =   1230
            Width           =   2895
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   8
            Left            =   2940
            MaxLength       =   10
            TabIndex        =   7
            Tag             =   "Cta.Cont.Impuesto|T|S|||sparam|ctaimpue|||"
            Top             =   1200
            Width           =   1215
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   10
            Left            =   2940
            MaxLength       =   10
            TabIndex        =   9
            Tag             =   "Raiz Cta.Socio|T|S|||sparam|raizctasoc|||"
            Top             =   1860
            Width           =   1215
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Enabled         =   0   'False
            Height          =   285
            Index           =   5
            Left            =   4200
            TabIndex        =   49
            Top             =   1860
            Width           =   2895
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   11
            Left            =   2940
            MaxLength       =   10
            TabIndex        =   10
            Tag             =   "Raiz Cta.Cliente|T|S|||sparam|raizctacli|||"
            Top             =   2190
            Width           =   1215
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Enabled         =   0   'False
            Height          =   285
            Index           =   6
            Left            =   4200
            TabIndex        =   48
            Top             =   2190
            Width           =   2895
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   9
            Left            =   2940
            MaxLength       =   10
            TabIndex        =   8
            Tag             =   "Cta.Cont.Fam.Dto|T|S|||sparam|ctafamdefecto|||"
            Top             =   1530
            Width           =   1215
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Enabled         =   0   'False
            Height          =   285
            Index           =   4
            Left            =   4200
            TabIndex        =   47
            Top             =   1530
            Width           =   2895
         End
         Begin VB.Label Label1 
            Caption         =   "Concepto al Haber Resto"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   14
            Left            =   540
            TabIndex        =   108
            Top             =   3150
            Width           =   1920
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   14
            Left            =   2565
            ToolTipText     =   "Buscar Concepto"
            Top             =   3150
            Width           =   240
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   10
            Left            =   2565
            ToolTipText     =   "Buscar Concepto"
            Top             =   2835
            Width           =   240
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   9
            Left            =   2565
            ToolTipText     =   "Buscar Concepto"
            Top             =   2520
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Concepto al Debe"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   5
            Left            =   540
            TabIndex        =   83
            Top             =   2520
            Width           =   1350
         End
         Begin VB.Label Label1 
            Caption         =   "Concepto al Haber Efectivo"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   4
            Left            =   540
            TabIndex        =   82
            Top             =   2835
            Width           =   2010
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   8
            Left            =   2565
            ToolTipText     =   "Buscar Diario"
            Top             =   3510
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Número Diario "
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   3
            Left            =   540
            TabIndex        =   81
            Top             =   3510
            Width           =   1350
         End
         Begin VB.Label Label1 
            Caption         =   "Cta.Contable Impuesto"
            Height          =   195
            Index           =   27
            Left            =   540
            TabIndex        =   60
            Top             =   1230
            Width           =   1980
         End
         Begin VB.Label Label1 
            Caption         =   "Cta.Diferencias Positivas"
            Height          =   195
            Index           =   26
            Left            =   540
            TabIndex        =   59
            Top             =   930
            Width           =   1950
         End
         Begin VB.Label Label1 
            Caption         =   "Cta.Diferencias Negativas"
            Height          =   195
            Index           =   25
            Left            =   540
            TabIndex        =   58
            Top             =   600
            Width           =   1920
         End
         Begin VB.Label Label1 
            Caption         =   "Cta.Contable Contado"
            Height          =   195
            Index           =   24
            Left            =   540
            TabIndex        =   57
            Top             =   270
            Width           =   1920
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   0
            Left            =   2580
            ToolTipText     =   "Buscar Cta.Contable"
            Top             =   240
            Width           =   240
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   1
            Left            =   2580
            ToolTipText     =   "Buscar Cta.Contable"
            Top             =   570
            Width           =   240
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   2
            Left            =   2580
            ToolTipText     =   "Buscar Cta.Contable"
            Top             =   900
            Width           =   240
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   3
            Left            =   2580
            ToolTipText     =   "Buscar Cta.Contable"
            Top             =   1230
            Width           =   240
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   5
            Left            =   2580
            ToolTipText     =   "Buscar Cta.Contable"
            Top             =   1860
            Width           =   240
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   6
            Left            =   2580
            ToolTipText     =   "Buscar Cta.Contable"
            Top             =   2190
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Raiz Cta.Contable Socio"
            Height          =   195
            Index           =   0
            Left            =   540
            TabIndex        =   56
            Top             =   1890
            Width           =   1845
         End
         Begin VB.Label Label1 
            Caption         =   "Raiz.Cta.Contable Cliente"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   1
            Left            =   540
            TabIndex        =   55
            Top             =   2190
            Width           =   1995
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   4
            Left            =   2580
            ToolTipText     =   "Buscar Cta.Contable"
            Top             =   1530
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Cta.Cble.Familias Defecto"
            Height          =   195
            Index           =   2
            Left            =   540
            TabIndex        =   54
            Top             =   1560
            Width           =   2070
         End
      End
      Begin VB.Frame Frame2 
         Height          =   1575
         Left            =   -74820
         TabIndex        =   40
         Top             =   345
         Width           =   7665
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   1
            Left            =   2550
            MaxLength       =   20
            TabIndex        =   0
            Tag             =   "Servidor Contabilidad|T|S|||sparam|serconta|||"
            Text            =   "3wwwwwwwwwwwwwwwwwwwwwwwwwwwww"
            Top             =   210
            Width           =   4560
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   0
            Left            =   4230
            MaxLength       =   15
            TabIndex        =   45
            Tag             =   "Código Parámetros Aplic|N|N|||sparam|codparam||S|"
            Text            =   "1"
            Top             =   240
            Width           =   645
         End
         Begin VB.TextBox Text1 
            Height          =   285
            IMEMode         =   3  'DISABLE
            Index           =   3
            Left            =   2550
            MaxLength       =   20
            PasswordChar    =   "*"
            TabIndex        =   2
            Tag             =   "Password Contabilidad|T|S|||sparam|pasconta|||"
            Text            =   "3"
            Top             =   840
            Width           =   4560
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   2
            Left            =   2550
            MaxLength       =   20
            TabIndex        =   1
            Tag             =   "Usuario Contabilidad|T|S|||sparam|usuconta|||"
            Text            =   "3wwwwwwwwwwwwwwwwwww"
            Top             =   525
            Width           =   4560
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   4
            Left            =   2550
            MaxLength       =   2
            TabIndex        =   3
            Tag             =   "Nº Contabilidad|N|S|||sparam|numconta|||"
            Text            =   "3"
            Top             =   1185
            Width           =   780
         End
         Begin VB.Label Label1 
            Caption         =   "Password"
            Height          =   195
            Index           =   15
            Left            =   510
            TabIndex        =   44
            Top             =   900
            Width           =   840
         End
         Begin VB.Label Label1 
            Caption         =   "Usuario"
            Height          =   195
            Index           =   17
            Left            =   510
            TabIndex        =   43
            Top             =   600
            Width           =   780
         End
         Begin VB.Label Label1 
            Caption         =   "Nº conta"
            Height          =   195
            Index           =   18
            Left            =   510
            TabIndex        =   42
            Top             =   1230
            Width           =   900
         End
         Begin VB.Label Label1 
            Caption         =   "Servidor"
            Height          =   195
            Index           =   19
            Left            =   510
            TabIndex        =   41
            Top             =   270
            Width           =   900
         End
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   34
         Left            =   -71085
         MaxLength       =   3
         TabIndex        =   127
         Tag             =   "Letra Serie|T|S|||sparam|letraint|||"
         Top             =   2850
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Tipo Precio"
         Height          =   195
         Index           =   9
         Left            =   -71010
         TabIndex        =   105
         Top             =   1605
         Width           =   2205
      End
      Begin VB.Label Label1 
         Caption         =   "Letra de Serie de Internas"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   8
         Left            =   -73320
         TabIndex        =   128
         Top             =   2880
         Width           =   1845
      End
      Begin VB.Image imgAyuda 
         Height          =   240
         Index           =   0
         Left            =   -70290
         MousePointer    =   4  'Icon
         Tag             =   "-1"
         ToolTipText     =   "Ayuda"
         Top             =   2880
         Width           =   240
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   7215
      TabIndex        =   34
      Top             =   6555
      Width           =   1035
   End
   Begin VB.Frame Frame1 
      Height          =   540
      Left            =   240
      TabIndex        =   37
      Top             =   6495
      Width           =   3000
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
         TabIndex        =   38
         Top             =   210
         Width           =   2760
      End
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   5940
      TabIndex        =   33
      Top             =   6555
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   7230
      TabIndex        =   35
      Top             =   6570
      Width           =   1035
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   36
      Top             =   0
      Width           =   8685
      _ExtentX        =   15319
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   5
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Añadir"
            Object.Tag             =   "1"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Modificar"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   330
      Left            =   3630
      Top             =   5250
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
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
   Begin VB.Menu mnOpciones 
      Caption         =   "&Opciones"
      Begin VB.Menu mnAñadir 
         Caption         =   "&Añadir"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnModificar 
         Caption         =   "&Modificar"
         HelpContextID   =   1
         Shortcut        =   ^M
      End
      Begin VB.Menu mnBarra1 
         Caption         =   "-"
      End
      Begin VB.Menu mnSalir 
         Caption         =   "&Salir"
         Shortcut        =   ^S
      End
   End
End
Attribute VB_Name = "frmConfParamAplic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ### [Monica] 06/09/2006
' procedimiento nuevo introducido de la gestion

Option Explicit

Private WithEvents frmCoop As frmManCoope
Attribute frmCoop.VB_VarHelpID = -1
Private WithEvents frmCtas As frmCtasConta
Attribute frmCtas.VB_VarHelpID = -1
Private WithEvents frmArt As frmManArtic
Attribute frmArt.VB_VarHelpID = -1
Private WithEvents frmTDia As frmDiaConta
Attribute frmTDia.VB_VarHelpID = -1
Private WithEvents frmConce As frmConceConta
Attribute frmConce.VB_VarHelpID = -1
Private WithEvents frmFam As frmManFamia
Attribute frmFam.VB_VarHelpID = -1
Private WithEvents frmIva As frmTipIVAConta
Attribute frmIva.VB_VarHelpID = -1


Private NombreTabla As String  'Nombre de la tabla o de la
Private Ordenacion As String
Private CadenaConsulta As String

Dim indice As Byte
Dim Encontrado As Boolean
Dim Modo As Byte
'0: Inicial
'2: Visualizacion
'3: Añadir
'4: Modificar

Dim indCodigo As Integer


Private Sub chkctrstock_GotFocus()
    PonerFocoChk chkctrstock
End Sub

Private Sub chkctrstock_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub chkOutlook_GotFocus()
    ConseguirfocoChk Modo
End Sub

Private Sub chkOutlook_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
      KEYpress KeyAscii
End Sub

Private Sub Combo1_GotFocus()
    If Modo = 1 Then Combo1.BackColor = vbYellow
End Sub

Private Sub Combo1_LostFocus()
    If Combo1.BackColor = vbYellow Then Combo1.BackColor = vbWhite
End Sub

Private Sub Combo2_KeyPress(KeyAscii As Integer)
      KEYpress KeyAscii
End Sub

Private Sub Combo2_GotFocus()
    If Modo = 1 Then Combo2.BackColor = vbYellow
End Sub

Private Sub Combo2_LostFocus()
    If Combo2.BackColor = vbYellow Then Combo2.BackColor = vbWhite
End Sub



Private Sub cmdAceptar_Click()
Dim actualiza As Boolean
Dim kms As Currency

    
    If Modo = 3 Then
        If DatosOk Then
            'Cambiamos el path
            'CambiaPath True
            If InsertarDesdeForm(Me) Then
                PonerModo 0
'                ActualizaNombreEmpresa
                MsgBox "Debe salir de la aplicacion para que los cambios tengan efecto", vbExclamation
            End If

        End If
    End If


    If Modo = 4 Then 'MODIFICAR
        If DatosOk Then
            If Not vParamAplic Is Nothing Then
                'Datos contabilidad
                vParamAplic.ServidorConta = Text1(1).Text
                vParamAplic.UsuarioConta = Text1(2).Text
                vParamAplic.PasswordConta = Text1(3).Text
                vParamAplic.NumeroConta = ComprobarCero(Text1(4).Text)
                
                vParamAplic.CtaContable = Text1(5).Text
                vParamAplic.CtaImpuesto = Text1(8).Text
                vParamAplic.CtaPositiva = Text1(7).Text
                vParamAplic.CtaNegativa = Text1(6).Text
                vParamAplic.TextoImpuesto = Text1(12).Text
                vParamAplic.Bonifact = Combo1.ListIndex
                vParamAplic.ArticDto = ComprobarCero(Text1(13).Text)
                vParamAplic.RaizCtaSoc = Text1(10).Text
                vParamAplic.RaizCtaCli = Text1(11).Text
                vParamAplic.CtaFamDefecto = Text1(9).Text
                vParamAplic.WebSoporte = Text1(14).Text
                vParamAplic.Cooperativa = Text1(15).Text
                vParamAplic.DireMail = Text1(16).Text
                vParamAplic.Smtphost = Text1(17).Text
                vParamAplic.SmtpUser = Text1(18).Text
                vParamAplic.Smtppass = Text1(19).Text
                vParamAplic.Cim = Text1(20).Text
                vParamAplic.Cee = Text1(21).Text
                
                
                vParamAplic.ConceptoDebe = ComprobarCero(Text1(22).Text)
                vParamAplic.ConceptoHaber = ComprobarCero(Text1(23).Text)
                vParamAplic.ConceptoHaberResto = ComprobarCero(Text1(37).Text)
                vParamAplic.NumDiario = ComprobarCero(Text1(24).Text)
                vParamAplic.ColecDefecto = ComprobarCero(Text1(25).Text)
                vParamAplic.FamDto = ComprobarCero(Text1(27).Text)
                vParamAplic.ImpresoraTarjetas = Replace(Text1(26).Text, "\", "\\")
                vParamAplic.ImpresoraTickets = Replace(Text1(32).Text, "\", "\\")
                vParamAplic.PathFacturaE = Replace(Text1(36).Text, "\", "\\")
                
                ' paramtros de compras
                vParamAplic.DiaPago1 = ComprobarCero(Text1(28).Text)
                vParamAplic.DiaPago2 = ComprobarCero(Text1(29).Text)
                vParamAplic.DiaPago3 = ComprobarCero(Text1(30).Text)
                
                vParamAplic.MesNoGirar = ComprobarCero(Text1(31).Text)
                vParamAplic.ControlStock = Me.chkctrstock.Value
                vParamAplic.TipoPrecio = Combo2.ListIndex
                

                ' FACTURAS INTERNAS
                vParamAplic.TipoIvaExento = ComprobarCero(Text1(33).Text)
                vParamAplic.LetraInt = Text1(34).Text
            
                vParamAplic.EnvioDesdeOutlook = Me.chkOutlook.Value
            
                ' Para utilizar el arigesmail
                vParamAplic.ExeEnvioMail = Trim(Text1(35).Text)

                ' Limite Fras efectivo
                vParamAplic.LimiteFra = ComprobarCero(Text1(38).Text)

                ' Tipos de iva
                vParamAplic.CodIvaGnral = ComprobarCero(Text1(39).Text)
                vParamAplic.CodIvaRedu = ComprobarCero(Text1(40).Text)
                vParamAplic.CodIvaSRedu = ComprobarCero(Text1(41).Text)

                vParamAplic.CodIvaGnralAnt = ComprobarCero(Text1(42).Text)
                vParamAplic.CodIvaReduAnt = ComprobarCero(Text1(43).Text)


                actualiza = vParamAplic.Modificar(Text1(0).Text)
                TerminaBloquear
    
                If actualiza Then  'Inserta o Modifica
                    'Abrir la conexion a la conta q hemos modificado
                    CerrarConexionConta
                    If vParamAplic.NumeroConta <> 0 Then
                        If Not AbrirConexionConta(vParamAplic.UsuarioConta, vParamAplic.PasswordConta) Then End
                    End If
                    PonerModo 2
                    PonerFocoBtn Me.cmdSalir
                End If
           End If
        End If
    End If
End Sub

Private Sub cmdCancelar_Click()
    TerminaBloquear
    If Data1.Recordset.EOF Then
        PonerModo 0
    Else
        PonerCampos
        PonerModo 2
    End If
End Sub

Private Sub cmdSalir_Click()
    Screen.MousePointer = vbDefault
    Unload Me
End Sub

Private Sub Form_Activate()
    Screen.MousePointer = vbDefault
    If Modo = 0 Then PonerCadenaBusqueda
    PonerFoco Text1(0)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub Form_Load()
Dim i As Byte
    'Icono del formulario
    Me.Icon = frmPpal.Icon

    ' ICONITOS DE LA BARRA
    With Me.Toolbar1
        .HotImageList = frmPpal.imgListComun_OM
        .DisabledImageList = frmPpal.imgListComun_BN
        .ImageList = frmPpal.imgListComun
        .Buttons(1).Image = 3   'Anyadir
        .Buttons(2).Image = 4   'Modificar
        .Buttons(5).Image = 11  'Salir
    End With
    
    imgAyuda(0).Picture = frmPpal.ImageListB.ListImages(10).Picture

    LimpiarCampos   'Limpia los campos TextBox
    CargaCombo
   
   'cargar IMAGES de busqueda
    For i = 0 To Me.imgBuscar.Count - 1
        Me.imgBuscar(i).Picture = frmPpal.imgListImages16.ListImages(1).Picture
    Next i

    SSTab1.Tab = 0

    NombreTabla = "sparam"
    Ordenacion = " ORDER BY codparam"
    'ASignamos un SQL al DATA1
    Data1.ConnectionString = Conn
    CadenaConsulta = "Select * from " & NombreTabla & Ordenacion
    Data1.RecordSource = CadenaConsulta
    Data1.Refresh
    
    Encontrado = True
    If Data1.Recordset.EOF Then
        'No hay registro de datos de parametros
        'quitar###
        Encontrado = False
    End If
    
    FrameCooperativa.Enabled = (vSesion.Nivel = 0)
    FrameCooperativa.visible = (vSesion.Nivel = 0)
    
    PonerModo 0

End Sub


Private Sub PonerCadenaBusqueda()
On Error GoTo EEPonerBusq
    Screen.MousePointer = vbHourglass

    Data1.RecordSource = CadenaConsulta
    Data1.Refresh
    If Data1.Recordset.RecordCount <= 0 Then
        LimpiarCampos
        MsgBox "No hay ningún registro en la tabla " & NombreTabla, vbInformation
        Screen.MousePointer = vbDefault
'        Me.Toolbar1.Buttons(1).Enabled = False 'Modificar
    Else
        Data1.Recordset.MoveFirst
        PonerCampos
    End If
    
    Screen.MousePointer = vbDefault
    Exit Sub
EEPonerBusq:
    MuestraError Err.Number, "PonerCadenaBusqueda"
    Screen.MousePointer = vbDefault
End Sub


Private Sub frmArt_DatoSeleccionado(CadenaSeleccion As String)
    Text1(13).Text = RecuperaValor(CadenaSeleccion, 1)
    FormateaCampo Text1(10)
    Text2(7).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmCoop_DatoSeleccionado(CadenaSeleccion As String)
'Cooperativa
    Text1(indice).Text = RecuperaValor(CadenaSeleccion, 1) 'codmacta
    Text2(indice).Text = RecuperaValor(CadenaSeleccion, 2) 'des macta
End Sub

Private Sub frmFam_DatoSeleccionado(CadenaSeleccion As String)
'Familia de descuento
    Text1(indice).Text = RecuperaValor(CadenaSeleccion, 1) 'codigo
    Text2(indice).Text = RecuperaValor(CadenaSeleccion, 2) 'desscripcion
End Sub

Private Sub frmCtas_DatoSeleccionado(CadenaSeleccion As String)
'Cuentas contables de la Contabilidad
    Text1(indice).Text = RecuperaValor(CadenaSeleccion, 1) 'codmacta
    Text2(indice - 5).Text = RecuperaValor(CadenaSeleccion, 2) 'des macta
End Sub

Private Sub frmConce_DatoSeleccionado(CadenaSeleccion As String)
'Cuentas contables de la Contabilidad
    Text1(indCodigo).Text = RecuperaValor(CadenaSeleccion, 1) 'codconce
    Text2(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2) 'des macta
End Sub

Private Sub frmIva_DatoSeleccionado(CadenaSeleccion As String)
'codigo de iva
    Text1(indCodigo).Text = RecuperaValor(CadenaSeleccion, 1) 'tipo de iva
    Text2(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2) 'descripcion
End Sub

Private Sub frmTDia_DatoSeleccionado(CadenaSeleccion As String)
'Cuentas contables de la Contabilidad
    Text1(indCodigo).Text = RecuperaValor(CadenaSeleccion, 1) 'diario
    Text2(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2) 'descripcion
End Sub


Private Sub imgAyuda_Click(Index As Integer)
Dim vCadena As String
    Select Case Index
        Case 0
           ' "____________________________________________________________"
            vCadena = "Necesitamos saber cual es la letra de serie de las Facturas Internas" & vbCrLf & _
                      "para distinguirlas de las que no lo son en la contabilización. " & vbCrLf & vbCrLf & _
                      "Las Facturas Internas cuando se contabilizan no van al registro" & vbCrLf & _
                      "de Iva de Clientes de Contabilidad, sino a un asiento del Diario." & vbCrLf
                      
                      
    End Select
    MsgBox vCadena, vbInformation, "Descripción de Ayuda"
    
End Sub

Private Sub imgBuscar_Click(Index As Integer)
Dim numNivel As Byte

    If vParamAplic.NumeroConta = 0 Then Exit Sub
    
    Select Case Index
        Case 0, 1, 2, 3, 4 'Cuentas Contables (de contabilidad)
            indice = Index + 5
            Set frmCtas = New frmCtasConta
            frmCtas.NumDigit = 0
            frmCtas.DatosADevolverBusqueda = "0|1|"
            frmCtas.CodigoActual = Text1(indice).Text
            frmCtas.Show vbModal
            Set frmCtas = Nothing
            PonerFoco Text1(indice)
                
        Case 5, 6 'raices de las cuentas contables de socio y cliente
            indice = Index + 5
            Set frmCtas = New frmCtasConta
            numNivel = DevuelveDesdeBDNew(cConta, "empresa", "numnivel", "codempre", Text1(4).Text, "N")
            frmCtas.NumDigit = DevuelveDesdeBDNew(cConta, "empresa", "numdigi" & numNivel - 1, "codempre", Text1(4).Text, "N")
            frmCtas.DatosADevolverBusqueda = "0|1|"
            frmCtas.CodigoActual = Text1(indice).Text
            frmCtas.Show vbModal
            Set frmCtas = Nothing
            If indice = 11 Then indice = 22
            PonerFoco Text1(indice)
        
        Case 7 'Articulo
            Set frmArt = New frmManArtic
            frmArt.DatosADevolverBusqueda = "0|1|"
            frmArt.CodigoActual = Text1(13).Text
            frmArt.Show vbModal
            Set frmArt = Nothing
            PonerFoco Text1(13)
            
        Case 9, 10 'conceptos
            AbrirFrmConceptos (Index + 13)
        
        Case 14 ' concepto
            AbrirFrmConceptos (Index + 23)
        
        Case 8 ' TIPOS DE DIARIO
            AbrirFrmDiario (Index + 16)
        
        Case 11 'Colectivo
            indice = Index + 14
            Set frmCoop = New frmManCoope
            frmCoop.DatosADevolverBusqueda = "0|1|"
            frmCoop.CodigoActual = Text1(25).Text
            frmCoop.Show vbModal
            Set frmCoop = Nothing
            PonerFoco Text1(25)
    
        Case 12 'Familia de Descuento
            indice = Index + 15
            Set frmFam = New frmManFamia
            frmFam.DatosADevolverBusqueda = "0|1|"
            frmFam.CodigoActual = Text1(25).Text
            frmFam.Show vbModal
            Set frmFam = Nothing
            PonerFoco Text1(25)
    
        'facturas internas
        Case 13 ' tipo de iva exento
            indCodigo = 33
            Set frmIva = New frmTipIVAConta
            frmIva.DatosADevolverBusqueda = "0|1|2|"
            frmIva.CodigoActual = Text1(indCodigo).Text
            frmIva.Show vbModal
            Set frmIva = Nothing
            PonerFoco Text1(indCodigo)
    
        ' tipos de iva actuales y anteriores
        Case 15, 16, 17, 18, 19
            indCodigo = Index + 24
            Set frmIva = New frmTipIVAConta
            frmIva.DatosADevolverBusqueda = "0|1|2|"
            frmIva.CodigoActual = Text1(indCodigo).Text
            frmIva.Show vbModal
            Set frmIva = Nothing
            PonerFoco Text1(indCodigo)
        
    End Select
End Sub


Private Sub mnAñadir_Click()
    If BLOQUEADesdeFormulario(Me) Then BotonAnyadir
End Sub

Private Sub mnModificar_Click()
    If BLOQUEADesdeFormulario(Me) Then BotonModificar
End Sub

Private Sub mnSalir_Click()
    Screen.MousePointer = vbDefault
    Unload Me
End Sub

Private Sub Text1_GotFocus(Index As Integer)
    ConseguirFoco Text1(Index), Modo
End Sub

Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'Avanzar/Retroceder los campos con las flechas de desplazamiento del teclado.
    KEYdown KeyCode
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
'14/02/2007 antes estaba esto
'    KEYpress (KeyAscii)
' ahora
    If KeyAscii = teclaBuscar Then
        Select Case Index
            Case 5: KEYBusqueda KeyAscii, 0 'cuenta contable contado
            Case 6: KEYBusqueda KeyAscii, 1 'cuenta de diferencias negativas
            Case 7: KEYBusqueda KeyAscii, 2 'cuenta de diferencias positivas
            Case 8: KEYBusqueda KeyAscii, 3 'cuenta contable impuesto
            Case 9: KEYBusqueda KeyAscii, 4 'cuenta contable familias defecto
            Case 10: KEYBusqueda KeyAscii, 5 'raiz cuenta contable socio
            Case 11: KEYBusqueda KeyAscii, 6 'raiz cuenta contable cliente
            Case 13: KEYBusqueda KeyAscii, 7 'articulo de descuento
            Case 22: KEYBusqueda KeyAscii, 9 'concepto al debe
            Case 23: KEYBusqueda KeyAscii, 10 'concepto al haber
            Case 37: KEYBusqueda KeyAscii, 14 'concepto al haber
            Case 24: KEYBusqueda KeyAscii, 8 'numero de diario
            Case 25: KEYBusqueda KeyAscii, 25 'cooperativa por defecto
            Case 27: KEYBusqueda KeyAscii, 27 'familia de descuento por defecto
            'facturas internas
            Case 33: KEYBusqueda KeyAscii, 13 'codigo de iva exento
        
            Case 39: KEYBusqueda KeyAscii, 39 'codigo de iva exento
            Case 40: KEYBusqueda KeyAscii, 40 'codigo de iva exento
            Case 41: KEYBusqueda KeyAscii, 41 'codigo de iva exento
            Case 42: KEYBusqueda KeyAscii, 42 'codigo de iva exento
            Case 43: KEYBusqueda KeyAscii, 43 'codigo de iva exento
        End Select
    Else
        KEYpress KeyAscii
    End If
End Sub

Private Sub KEYBusqueda(KeyAscii As Integer, indice As Integer)
    KeyAscii = 0
    imgBuscar_Click (indice)
End Sub

Private Sub Text1_LostFocus(Index As Integer)
Dim cadMen As String

'    If Text1(Index).Text = "" Then Exit Sub

    'Quitar espacios en blanco
    Text1(Index).Text = Trim(Text1(Index).Text)
    
    Select Case Index
        Case 4, 15 'numero de contabilidad y cooperativa
            If Not EsNumerico(Text1(Index).Text) Then
                Text1(Index).Text = ""
                PonerFoco Text1(Index)
            End If
            
        Case 5, 6, 7, 8, 9
            If Text1(Index).Text = "" Then Exit Sub
            Text2(Index - 5).Text = PonerNombreCuenta(Text1(Index), Modo)
            
        Case 10, 11
            If Text1(Index).Text = "" Then Exit Sub
            Text2(Index - 5).Text = NombreCuentaCorrecta(Text1(Index).Text)
            If Index = 11 Then
                SSTab1.Tab = 1
                PonerFoco Text1(12)
            End If
        
        Case 13 ' codigo de articulo
            If PonerFormatoEntero(Text1(Index)) Then
                Text2(7).Text = PonerNombreDeCod(Text1(Index), "sartic", "nomartic", "codartic", "N")
                If Text2(7).Text = "" Then
                    cadMen = "No existe el Articulo: " & Text1(Index).Text & vbCrLf
                    cadMen = cadMen & "¿Desea crearlo?" & vbCrLf
                    If MsgBox(cadMen, vbQuestion + vbYesNo) = vbYes Then
                        Set frmArt = New frmManArtic
                        frmArt.DatosADevolverBusqueda = "0|1|"
                        frmArt.NuevoCodigo = Text1(Index).Text
                        Text1(Index).Text = ""
                        TerminaBloquear
                        frmArt.Show vbModal
                        Set frmArt = Nothing
                        If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Data1, 1
                    Else
                        Text1(Index).Text = ""
                    End If
                    PonerFoco Text1(Index)
                End If
            End If
        ' conceptos al debe y al haber
        Case 22, 23, 37
            If Text1(Index).Text = "" Then Exit Sub
            Text2(Index).Text = PonerNombreConcepto(Text1(Index))
        ' numero de diario
        Case 24
            If Text1(Index).Text = "" Then Exit Sub
            Text2(Index).Text = DevuelveDesdeBDNew(cConta, "tiposdiario", "desdiari", "numdiari", Text1(24).Text, "N")
            
        Case 25 ' codigo de colectivo
            If PonerFormatoEntero(Text1(Index)) Then
                Text2(25).Text = PonerNombreDeCod(Text1(Index), "scoope", "nomcoope", "codcoope", "N")
                If Text2(25).Text = "" Then
                    cadMen = "No existe el Colectivo: " & Text1(Index).Text & vbCrLf
                    cadMen = cadMen & "¿Desea crearlo?" & vbCrLf
                    If MsgBox(cadMen, vbQuestion + vbYesNo) = vbYes Then
                        Set frmCoop = New frmManCoope
                        frmCoop.DatosADevolverBusqueda = "0|1|"
                        frmCoop.NuevoCodigo = Text1(Index).Text
                        Text1(Index).Text = ""
                        TerminaBloquear
                        frmCoop.Show vbModal
                        Set frmCoop = Nothing
                        If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Data1, 1
                    Else
                        Text1(Index).Text = ""
                    End If
                    PonerFoco Text1(Index)
                End If
            End If
        Case 27 ' familia de descuento
            If PonerFormatoEntero(Text1(Index)) Then
                Text2(27).Text = PonerNombreDeCod(Text1(Index), "sfamia", "nomfamia", "codfamia", "N")
                If Text2(27).Text = "" Then
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
            End If
            
        ' FACTURAS INTERNAS
        Case 33 ' codigo iva exento
            If PonerFormatoEntero(Text1(Index)) Then
                Text2(Index).Text = DevuelveDesdeBDNew(cConta, "tiposiva", "nombriva", "codigiva", Text1(Index), "N")
            Else
                Text2(Index).Text = ""
            End If
            
        Case 38 ' importe limite de facturas de efectivo
            PonerFormatoDecimal Text1(38), 3
    
    
        Case 39, 40, 41, 42, 43 ' codigo iva exento
            If PonerFormatoEntero(Text1(Index)) Then
                Text2(Index).Text = DevuelveDesdeBDNew(cConta, "tiposiva", "nombriva", "codigiva", Text1(Index), "N")
            Else
                Text2(Index).Text = ""
            End If
    
    
    End Select
        
End Sub


Private Sub Text1_Validate(Index As Integer, Cancel As Boolean)
    Select Case Index
        Case 6, 7
            If Text1(Index).Text <> "" Then
                If Not EsNumerico(Text1(Index).Text) Then
                    Cancel = True
                    ConseguirFoco Text1(Index), Modo
                End If
            End If
    End Select
    
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1  'Anyadir
            BotonAnyadir
        Case 2  'Modificar
            mnModificar_Click
        Case 5 'Salir
            mnSalir_Click
    End Select
End Sub


Private Sub BotonAnyadir()
    LimpiarCampos
    PonerModo 3
    Text1(0).Text = 1
    PonerFoco Text1(1)
End Sub


Private Sub BotonModificar()
    PonerModo 4
    PonerFoco Text1(1)
End Sub


Private Function DatosOk() As Boolean
Dim b As Boolean
    DatosOk = False
    b = CompForm(Me)
    DatosOk = b
End Function


Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, Modo, cerrar
    If cerrar Then Unload Me
End Sub


Private Sub PonerBotonCabecera(b As Boolean)
    Me.cmdAceptar.visible = Not b
    Me.cmdCancelar.visible = Not b
    Me.cmdSalir.visible = b
'    If b Then Me.lblIndicador.Caption = ""
End Sub


Private Sub PonerCampos()
Dim i As Byte

On Error GoTo EPonerCampos

    If Data1.Recordset.EOF Then Exit Sub
    PonerCamposForma Me, Data1
    
    ' ************* configurar els camps de les descripcions de les comptes *************
    For i = 0 To 4
        Text2(i).Text = PonerNombreCuenta(Text1(i + 5), Modo)
    Next i
    For i = 5 To 6
        Text2(i).Text = NombreCuentaCorrecta(Text1(i + 5).Text)
    Next i
    
    ' mostramos el nombre del articulo
    If Text1(13).Text <> "" Then
        Text2(7).Text = DevuelveDesdeBD("nomartic", "sartic", "codartic", Text1(13).Text, "N")
    End If
    Text1(15).Text = Data1.Recordset!Cooperativa
    ' ********************************************************************************
    ' numero de conceptos
    For i = 22 To 23
        Text2(i).Text = PonerNombreConcepto(Text1(i))
    Next i
    For i = 37 To 37
        Text2(i).Text = PonerNombreConcepto(Text1(i))
    Next i
    
    ' numero de diario
    Text2(24).Text = DevuelveDesdeBDNew(cConta, "tiposdiario", "desdiari", "numdiari", Text1(24).Text, "N")
    
    ' colectivo por defecto
    If Text1(25).Text <> "" Then
        Text2(25).Text = DevuelveDesdeBD("nomcoope", "scoope", "codcoope", Text1(25).Text, "N")
    End If
    
    ' familia de descuento
    If Text1(27).Text <> "" Then
        Text2(27).Text = DevuelveDesdeBD("nomfamia", "sfamia", "codfamia", Text1(27).Text, "N")
    End If
    If Text1(33).Text <> "" Then
        Text2(33).Text = DevuelveDesdeBDNew(cConta, "tiposiva", "nombriva", "codigiva", Text1(33), "N")
    End If
    
    ' tipos de iva actuales y antiguos
    'gnral
    If Text1(39).Text <> "" Then
        Text2(39).Text = DevuelveDesdeBDNew(cConta, "tiposiva", "nombriva", "codigiva", Text1(39), "N")
    End If
    'reducido
    If Text1(40).Text <> "" Then
        Text2(40).Text = DevuelveDesdeBDNew(cConta, "tiposiva", "nombriva", "codigiva", Text1(40), "N")
    End If
    'super-reducido
    If Text1(41).Text <> "" Then
        Text2(41).Text = DevuelveDesdeBDNew(cConta, "tiposiva", "nombriva", "codigiva", Text1(41), "N")
    End If
    'gnral antiguo
    If Text1(42).Text <> "" Then
        Text2(42).Text = DevuelveDesdeBDNew(cConta, "tiposiva", "nombriva", "codigiva", Text1(42), "N")
    End If
    'reducido antiguo
    If Text1(43).Text <> "" Then
        Text2(43).Text = DevuelveDesdeBDNew(cConta, "tiposiva", "nombriva", "codigiva", Text1(43), "N")
    End If
    
EPonerCampos:
    If Err.Number <> 0 Then MuestraError Err.Number, "Poniendo Campos", Err.Description
End Sub


Private Sub PonerOpcionesMenu()
    PonerOpcionesMenuGeneral Me
End Sub


Private Sub LimpiarCampos()
    limpiar Me   'Metodo general: Limpia los controles TextBox
    lblIndicador.Caption = ""
    'Aqui va el especifico de cada form es
    '### a mano
End Sub


'----------------------------------------------------------------
'   En PONERMODO se habilitan, o no, los diverso campos del
'   formulario en funcion del modo en k vayamos a trabajar
Private Sub PonerModo(Kmodo As Byte)
Dim b As Boolean
Dim i As Byte

    Modo = Kmodo
    PonerIndicador lblIndicador, Modo
      
    '------------------------------------------------------
    'Modo insertar o modificar
    b = (Kmodo >= 3) '-->Luego not b sera kmodo<3
    If cmdCancelar.visible Then
        cmdCancelar.Cancel = True
    Else
        cmdCancelar.Cancel = False
    End If
    PonerBotonCabecera Not b
       
    'Bloquea los campos Text1 sino estamos modificando/Insertando Datos
    'Si estamos en Insertar además limpia los campos Text1 y bloquea la clave primaria
    BloquearText1 Me, Modo
    
    BloquearChecks Me, Modo
    
    
    'Bloquear el combobox
    Me.Combo1.Enabled = (Modo >= 3)
    BloquearCmb Combo1, Not (Modo >= 3)
    
    Me.Combo2.Enabled = (Modo >= 3)
    BloquearCmb Combo2, Not (Modo >= 3)
    
    'Bloquear imagen de Busqueda
    For i = 0 To Me.imgBuscar.Count - 1
        Me.imgBuscar(i).Enabled = (Modo >= 3)
    Next i

    BloquearImgBuscar Me, Modo

    PonerModoOpcionesMenu 'Activar opciones de menu según el Modo
    PonerOpcionesMenu   'Activar opciones de menu según nivel
                        'de permisos del usuario
End Sub


Private Sub PonerModoOpcionesMenu()
Dim b As Boolean
    b = (Modo = 3) Or (Modo = 4)
    Me.Toolbar1.Buttons(1).Enabled = Not Encontrado And Not b  'Añadir
    Me.Toolbar1.Buttons(2).Enabled = Encontrado And Not b 'Modificar
    Me.mnAñadir.Enabled = Not Encontrado And Not b
    Me.mnModificar.Enabled = Encontrado And Not b
'    Me.Toolbar1.Buttons(2).Enabled = (Not b) 'Modificar
End Sub


Private Sub CargaCombo()
Dim Ini As Integer
Dim Fin As Integer
Dim i As Integer

    ' *** neteje els combos, els pose valor i seleccione el valor per defecte ***
    Combo1.Clear
    
    Combo1.AddItem "No"
    Combo1.ItemData(Combo1.NewIndex) = 0
    Combo1.AddItem "Si"
    Combo1.ItemData(Combo1.NewIndex) = 1


    Combo2.Clear
    
    Combo2.AddItem "Precio Medio Ponderado"
    Combo2.ItemData(Combo2.NewIndex) = 0
    Combo2.AddItem "Ultimo Precio de Compra"
    Combo2.ItemData(Combo2.NewIndex) = 1



End Sub

Private Sub AbrirFrmDiario(indice As Integer)
    indCodigo = indice
    Set frmTDia = New frmDiaConta
    frmTDia.DatosADevolverBusqueda = "0|1|"
    frmTDia.CodigoActual = Text2(indCodigo)
    frmTDia.Show vbModal
    Set frmTDia = Nothing
End Sub

Private Sub AbrirFrmConceptos(indice As Integer)
    indCodigo = indice
    Set frmConce = New frmConceConta
    frmConce.DatosADevolverBusqueda = "0|1|"
    frmConce.CodigoActual = Text1(indCodigo)
    frmConce.Show vbModal
    Set frmConce = Nothing
End Sub
 

