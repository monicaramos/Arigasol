VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmHcoFact 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Histórico de Facturas"
   ClientHeight    =   7995
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11835
   Icon            =   "frmHcoFact.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7995
   ScaleWidth      =   11835
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   1995
      Left            =   240
      TabIndex        =   46
      Top             =   2130
      Width           =   11385
      _ExtentX        =   20082
      _ExtentY        =   3519
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   6
      TabHeight       =   520
      ForeColor       =   9907723
      TabCaption(0)   =   "Total Factura"
      TabPicture(0)   =   "frmHcoFact.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "FrameTotFactu"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Rectificativa"
      TabPicture(1)   =   "frmHcoFact.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "FrameRectificativa"
      Tab(1).ControlCount=   1
      Begin VB.Frame FrameRectificativa 
         BorderStyle     =   0  'None
         Caption         =   "Datos Factura que Rectifica"
         ForeColor       =   &H00972E0B&
         Height          =   1455
         Left            =   -74970
         TabIndex        =   70
         Top             =   330
         Width           =   11235
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   21
            Left            =   1440
            MaxLength       =   7
            TabIndex        =   74
            Tag             =   "Nº de Factura rectifica|N|S|||schfac|rectif_numfactu|0000000||"
            Top             =   705
            Width           =   795
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   22
            Left            =   1440
            MaxLength       =   3
            TabIndex        =   73
            Tag             =   "Letra Serie|T|S|||schfac|rectif_letraser|||"
            Top             =   360
            Width           =   495
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   23
            Left            =   1440
            MaxLength       =   10
            TabIndex        =   72
            Tag             =   "Fecha Factura Rectif|F|S|||schfac|rectif_fecfactu|dd/mm/yyyy||"
            Top             =   1080
            Width           =   1035
         End
         Begin VB.TextBox Text1 
            Height          =   855
            Index           =   24
            Left            =   4920
            MaxLength       =   100
            TabIndex        =   71
            Tag             =   "Motivo Rectificativa|T|S|||schfac|observac|||"
            Top             =   540
            Width           =   6255
         End
         Begin VB.Label Label1 
            Caption         =   "Nº Factura"
            Height          =   255
            Index           =   7
            Left            =   240
            TabIndex        =   78
            Top             =   735
            Width           =   855
         End
         Begin VB.Image imgFec 
            Height          =   240
            Index           =   23
            Left            =   1125
            Picture         =   "frmHcoFact.frx":0044
            ToolTipText     =   "Buscar fecha"
            Top             =   1080
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Letra Serie"
            Height          =   255
            Index           =   9
            Left            =   240
            TabIndex        =   77
            Top             =   360
            Width           =   855
         End
         Begin VB.Label Label1 
            Caption         =   "Fecha Fact."
            Height          =   255
            Index           =   12
            Left            =   240
            TabIndex        =   76
            Top             =   1080
            Width           =   855
         End
         Begin VB.Label Label1 
            Caption         =   "Motivo"
            Height          =   255
            Index           =   17
            Left            =   4920
            TabIndex        =   75
            Top             =   240
            Width           =   1155
         End
      End
      Begin VB.Frame FrameTotFactu 
         BorderStyle     =   0  'None
         ForeColor       =   &H00972E0B&
         Height          =   1545
         Left            =   60
         TabIndex        =   47
         Top             =   330
         Width           =   11235
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0C0FF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   25
            Left            =   5760
            MaxLength       =   15
            TabIndex        =   79
            Top             =   1140
            Width           =   2325
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   14
            Left            =   240
            MaxLength       =   15
            TabIndex        =   62
            Tag             =   "Base IVA 3|N|S|||schfac|baseimp3|#,###,###,##0.00|N|"
            Top             =   1185
            Width           =   1605
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   10
            Left            =   240
            MaxLength       =   15
            TabIndex        =   61
            Tag             =   "Base IVA 2|N|S|||schfac|baseimp2|#,###,###,##0.00|N|"
            Top             =   840
            Width           =   1605
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   6
            Left            =   240
            MaxLength       =   15
            TabIndex        =   60
            Tag             =   "Base IVA 1|N|N|||schfac|baseimp1|#,###,###,##0.00|N|"
            Text            =   "575757575757557"
            Top             =   480
            Width           =   1605
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   17
            Left            =   4000
            MaxLength       =   15
            TabIndex        =   59
            Tag             =   "Importe IVA 3|N|S|||schfac|impoiva3|#,###,###,##0.00|N|"
            Top             =   1185
            Width           =   1605
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   13
            Left            =   4000
            MaxLength       =   15
            TabIndex        =   58
            Tag             =   "Importe IVA 2|N|S|||schfac|impoiva2|#,###,###,##0.00|N|"
            Top             =   840
            Width           =   1605
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   9
            Left            =   4000
            MaxLength       =   15
            TabIndex        =   57
            Tag             =   "Importe IVA 1|N|S|||schfac|impoiva1|#,###,###,##0.00|N|"
            Top             =   510
            Width           =   1605
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   16
            Left            =   3210
            MaxLength       =   6
            TabIndex        =   56
            Tag             =   "% IVA 3|N|S|0|100.00|schfac|porciva3|##0.00|N|"
            Top             =   1185
            Width           =   645
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   12
            Left            =   3210
            MaxLength       =   6
            TabIndex        =   55
            Tag             =   "% IVA 2|N|S|0|100.00|schfac|porciva2|##0.00|N|"
            Top             =   870
            Width           =   645
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   8
            Left            =   3210
            MaxLength       =   6
            TabIndex        =   54
            Tag             =   "% IVA 1|N|S|0|100.00|schfac|porciva1|##0.00|N|"
            Text            =   "99.99"
            Top             =   510
            Width           =   645
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   15
            Left            =   2540
            MaxLength       =   2
            TabIndex        =   53
            Tag             =   "Tipo IVA 3|N|S|0|99|schfac|tipoiva3|00||"
            Top             =   1185
            Width           =   525
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   11
            Left            =   2540
            MaxLength       =   2
            TabIndex        =   52
            Tag             =   "Tipo IVA 2|N|S|0|99|schfac|tipoiva2|00||"
            Top             =   840
            Width           =   525
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   7
            Left            =   2540
            MaxLength       =   2
            TabIndex        =   51
            Tag             =   "Tipo IVA 1|N|N|0|99|schfac|tipoiva1|00||"
            Text            =   "12"
            Top             =   510
            Width           =   525
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00CAE3FD&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   18
            Left            =   5760
            MaxLength       =   15
            TabIndex        =   50
            Tag             =   "Total Factura|N|N|-9999999999.99|9999999999.99|schfac|totalfac|#,###,###,##0.00|N|"
            Top             =   510
            Width           =   2325
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0C0FF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   19
            Left            =   8280
            MaxLength       =   15
            TabIndex        =   49
            Tag             =   "Impuesto|N|N|-9999999999.99|9999999999.99|schfac|impuesto|#,###,###,##0.00|N|"
            Top             =   510
            Width           =   2325
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0C0FF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   20
            Left            =   8280
            MaxLength       =   15
            TabIndex        =   48
            Tag             =   "Imp.Sigaus|N|S|-9999999999.99|9999999999.99|schfac|impuesigaus|#,###,###,##0.00|N|"
            Top             =   1140
            Width           =   2325
         End
         Begin VB.Label Label1 
            Caption         =   "Importe Vale"
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
            Index           =   18
            Left            =   5760
            TabIndex        =   80
            Top             =   870
            Width           =   2205
         End
         Begin VB.Label Label1 
            Caption         =   "Base Imponible"
            Height          =   255
            Index           =   10
            Left            =   240
            TabIndex        =   69
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label Label1 
            Caption         =   "Importe IVA"
            Height          =   255
            Index           =   16
            Left            =   4000
            TabIndex        =   68
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label Label1 
            Caption         =   "% IVA"
            Height          =   255
            Index           =   15
            Left            =   3240
            TabIndex        =   67
            Top             =   240
            Width           =   615
         End
         Begin VB.Label Label1 
            Caption         =   "Tipo IVA"
            Height          =   255
            Index           =   14
            Left            =   2540
            TabIndex        =   66
            Top             =   240
            Width           =   615
         End
         Begin VB.Label Label1 
            Caption         =   "Total Factura"
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
            Index           =   11
            Left            =   5760
            TabIndex        =   65
            Top             =   240
            Width           =   1575
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   3
            Left            =   2240
            MousePointer    =   4  'Icon
            Tag             =   "-1"
            ToolTipText     =   "Buscar tipo de IVA"
            Top             =   510
            Width           =   240
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   4
            Left            =   2240
            MousePointer    =   4  'Icon
            Tag             =   "-1"
            ToolTipText     =   "Buscar tipo de IVA"
            Top             =   840
            Width           =   240
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   5
            Left            =   2240
            MousePointer    =   4  'Icon
            Tag             =   "-1"
            ToolTipText     =   "Buscar tipo de IVA"
            Top             =   1200
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Impuesto"
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
            Index           =   8
            Left            =   8280
            TabIndex        =   64
            Top             =   240
            Width           =   1575
         End
         Begin VB.Label Label1 
            Caption         =   "Impuesto Sigaus"
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
            Index           =   0
            Left            =   8280
            TabIndex        =   63
            Top             =   870
            Width           =   2205
         End
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1605
      Index           =   0
      Left            =   240
      TabIndex        =   25
      Top             =   480
      Width           =   11415
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   26
         Left            =   7185
         TabIndex        =   81
         Top             =   1260
         Width           =   3975
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   26
         Left            =   6285
         MaxLength       =   4
         TabIndex        =   19
         Tag             =   "Departamento|N|N|0|9999|schfac|coddepar|0000||"
         Top             =   1245
         Width           =   675
      End
      Begin VB.CheckBox Check1 
         Caption         =   "FacturaE"
         Height          =   255
         Index           =   0
         Left            =   2970
         TabIndex        =   45
         Tag             =   "FacturaE|N|N|0|1|schfac|enfacturae|||"
         Top             =   555
         Width           =   1275
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Contabilizada"
         Height          =   255
         Index           =   1
         Left            =   2970
         TabIndex        =   44
         Tag             =   "Contabilizada|N|N|0|1|schfac|intconta|||"
         Top             =   270
         Width           =   1425
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   4
         Left            =   6285
         MaxLength       =   2
         TabIndex        =   17
         Tag             =   "Colectivo|N|N|0|99|schfac|codcoope|00||"
         Top             =   550
         Width           =   360
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   5
         Left            =   6285
         MaxLength       =   2
         TabIndex        =   18
         Tag             =   "Forma de pago|N|N|0|99|schfac|codforpa|00||"
         Top             =   900
         Width           =   360
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   5
         Left            =   7200
         TabIndex        =   32
         Top             =   900
         Width           =   3975
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   1
         Left            =   1440
         MaxLength       =   7
         TabIndex        =   14
         Tag             =   "Nº de Factura|N|N|0|9999999|schfac|numfactu|0000000|S|"
         Top             =   690
         Width           =   795
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   3
         Left            =   6285
         MaxLength       =   6
         TabIndex        =   16
         Tag             =   "Cliente|N|N|0|999999|schfac|codsocio|000000|N|"
         Top             =   200
         Width           =   720
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   3
         Left            =   7200
         TabIndex        =   30
         Top             =   200
         Width           =   3975
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   4
         Left            =   7200
         TabIndex        =   29
         Top             =   550
         Width           =   3975
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   0
         Left            =   1440
         MaxLength       =   3
         TabIndex        =   13
         Tag             =   "Letra Serie|T|N|||schfac|letraser||S|"
         Top             =   285
         Width           =   495
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   2
         Left            =   1440
         MaxLength       =   10
         TabIndex        =   15
         Tag             =   "Fecha Factura|F|N|||schfac|fecfactu|dd/mm/yyyy|S|"
         Top             =   1095
         Width           =   1035
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   6
         Left            =   6015
         Tag             =   "-1"
         ToolTipText     =   "Buscar Departamento"
         Top             =   1260
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Departamento"
         Height          =   255
         Index           =   19
         Left            =   4905
         TabIndex        =   82
         Top             =   1260
         Width           =   1035
      End
      Begin VB.Label Label1 
         Caption         =   "Cliente"
         Height          =   255
         Index           =   13
         Left            =   4920
         TabIndex        =   43
         Top             =   195
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Forma pago"
         Height          =   255
         Index           =   5
         Left            =   4920
         TabIndex        =   33
         Top             =   900
         Width           =   855
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   2
         Left            =   6015
         Tag             =   "-1"
         ToolTipText     =   "Buscar Forma de Pago"
         Top             =   900
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Nº Factura"
         Height          =   255
         Index           =   4
         Left            =   210
         TabIndex        =   31
         Top             =   690
         Width           =   855
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   0
         Left            =   6015
         MousePointer    =   4  'Icon
         Tag             =   "-1"
         ToolTipText     =   "Buscar Cliente"
         Top             =   180
         Width           =   240
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   2
         Left            =   1140
         Picture         =   "frmHcoFact.frx":00CF
         ToolTipText     =   "Buscar fecha"
         Top             =   1095
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   1
         Left            =   6015
         Tag             =   "-1"
         ToolTipText     =   "Buscar Colectivo"
         Top             =   555
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Letra Serie"
         Height          =   255
         Index           =   2
         Left            =   210
         TabIndex        =   28
         Top             =   285
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha Fact."
         Height          =   255
         Index           =   1
         Left            =   210
         TabIndex        =   27
         Top             =   1095
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Colectivo"
         Height          =   255
         Index           =   3
         Left            =   4920
         TabIndex        =   26
         Top             =   555
         Width           =   855
      End
   End
   Begin VB.Frame FrameAux0 
      Caption         =   "Lineas Factura"
      ForeColor       =   &H00972E0B&
      Height          =   3255
      Left            =   240
      TabIndex        =   34
      Top             =   4170
      Width           =   11415
      Begin VB.TextBox txtAux 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   290
         Index           =   11
         Left            =   9600
         MaxLength       =   10
         TabIndex        =   11
         Tag             =   "Precio|N|N|||slhfac|preciove|#,##0.000||"
         Text            =   "Pre"
         Top             =   1920
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox txtAux 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   290
         Index           =   10
         Left            =   8640
         MaxLength       =   10
         TabIndex        =   10
         Tag             =   "Cantidad|N|N|||slhfac|cantidad|#,##0.00||"
         Text            =   "Can"
         Top             =   1920
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox txtAux 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   290
         Index           =   8
         Left            =   4920
         MaxLength       =   13
         TabIndex        =   8
         Tag             =   "Tarjeta|N|N|0|9999999999999|slhfac|numtarje|0000000000000||"
         Text            =   "tar"
         Top             =   1920
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox txtAux 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   290
         Index           =   7
         Left            =   4560
         MaxLength       =   1
         TabIndex        =   7
         Tag             =   "Turno|N|N|0|9|slhfac|codturno|0||"
         Text            =   "T"
         Top             =   1920
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.TextBox txtAux 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   290
         Index           =   6
         Left            =   3960
         MaxLength       =   5
         TabIndex        =   6
         Tag             =   "Hora|FHH|N|||slhfac|horalbar|hh:mm||"
         Text            =   "Hor"
         Top             =   1920
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox txtAux 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   290
         Index           =   5
         Left            =   3000
         MaxLength       =   10
         TabIndex        =   5
         Tag             =   "Fecha Albaran|F|N|||slhfac|fecalbar|dd/mm/yyyy||"
         Text            =   "Fec"
         Top             =   1920
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.CommandButton btnBuscar 
         Appearance      =   0  'Flat
         Caption         =   "+"
         Height          =   300
         Index           =   0
         Left            =   6120
         MaskColor       =   &H00000000&
         TabIndex        =   39
         ToolTipText     =   "Buscar Artículo"
         Top             =   1920
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.TextBox txtAux 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   290
         Index           =   4
         Left            =   2040
         MaxLength       =   8
         TabIndex        =   4
         Tag             =   "Albaran|T|N|||slhfac|numalbar|||"
         Text            =   "Alb"
         Top             =   1920
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox txtAux 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   290
         Index           =   9
         Left            =   5640
         MaxLength       =   6
         TabIndex        =   9
         Tag             =   "Artículo|N|N|0|999999|slhfac|codartic|000000||"
         Text            =   "Arti"
         Top             =   1920
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox txtAux 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   290
         Index           =   0
         Left            =   120
         MaxLength       =   3
         TabIndex        =   0
         Tag             =   "Letra Serie|T|N|||slhfac|letraser||S|"
         Text            =   "L"
         Top             =   1920
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox txtAux 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   290
         Index           =   3
         Left            =   1680
         MaxLength       =   2
         TabIndex        =   3
         Tag             =   "Número de línea|N|N|1|9999|slhfac|numlinea|0000|S|"
         Text            =   "li"
         Top             =   1920
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.TextBox txtAux 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   290
         Index           =   1
         Left            =   480
         MaxLength       =   7
         TabIndex        =   1
         Tag             =   "Nº Factura|N|N|0|9999999|slhfac|numfactu|0000000|S|"
         Text            =   "Fac"
         Top             =   1920
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox txtAux 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   290
         Index           =   2
         Left            =   1080
         MaxLength       =   10
         TabIndex        =   2
         Tag             =   "Fecha Factura|F|N|||slhfac|fecfactu|dd/mm/yyyy|S|"
         Text            =   "fecfactu"
         Top             =   1920
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox txtAux 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   290
         Index           =   12
         Left            =   10440
         MaxLength       =   15
         TabIndex        =   12
         Tag             =   "Importe|N|N|||slhfac|implinea|##,###,##0.00||"
         Text            =   "Importe"
         Top             =   1920
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox txtAux2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   6360
         TabIndex        =   35
         Top             =   1920
         Visible         =   0   'False
         Width           =   2175
      End
      Begin MSAdodcLib.Adodc AdoAux 
         Height          =   375
         Index           =   0
         Left            =   4560
         Top             =   240
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
      Begin MSComctlLib.Toolbar ToolAux 
         Height          =   390
         Index           =   0
         Left            =   240
         TabIndex        =   36
         Top             =   270
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
               Enabled         =   0   'False
               Object.Visible         =   0   'False
               Object.Tag             =   "2"
               Style           =   3
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
         Begin VB.CheckBox Check2 
            Caption         =   "Vista previa"
            Height          =   195
            Index           =   1
            Left            =   8400
            TabIndex        =   37
            Top             =   120
            Visible         =   0   'False
            Width           =   1215
         End
      End
      Begin MSDataGridLib.DataGrid DataGridAux 
         Height          =   2310
         Index           =   0
         Left            =   240
         TabIndex        =   38
         Top             =   735
         Width           =   10755
         _ExtentX        =   18971
         _ExtentY        =   4075
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
   Begin VB.Frame Frame1 
      Height          =   555
      Index           =   0
      Left            =   240
      TabIndex        =   23
      Top             =   7395
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
         TabIndex        =   24
         Top             =   180
         Width           =   2655
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   10440
      TabIndex        =   21
      Top             =   7515
      Width           =   1035
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   9120
      TabIndex        =   20
      Top             =   7500
      Width           =   1035
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   375
      Left            =   4200
      Top             =   6960
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
   Begin VB.CommandButton cmdRegresar 
      Caption         =   "&Regresar"
      Height          =   375
      Left            =   10440
      TabIndex        =   22
      Top             =   7500
      Visible         =   0   'False
      Width           =   1035
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   41
      Top             =   0
      Width           =   11835
      _ExtentX        =   20876
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
            Object.ToolTipText     =   "Rectificar factura"
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
         TabIndex        =   42
         Top             =   120
         Width           =   1215
      End
   End
   Begin VB.Label Label1 
      Caption         =   "Empresa"
      Height          =   255
      Index           =   6
      Left            =   360
      TabIndex        =   40
      Top             =   720
      Width           =   615
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
         Enabled         =   0   'False
         HelpContextID   =   2
         Shortcut        =   ^N
         Visible         =   0   'False
      End
      Begin VB.Menu mnModificar 
         Caption         =   "&Modificar"
         HelpContextID   =   2
         Shortcut        =   ^M
      End
      Begin VB.Menu mnEliminar 
         Caption         =   "&Eliminar"
         Enabled         =   0   'False
         HelpContextID   =   2
         Shortcut        =   ^E
         Visible         =   0   'False
      End
      Begin VB.Menu mnRectificar 
         Caption         =   "&Rectificar"
         Shortcut        =   ^R
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
Attribute VB_Name = "frmHcoFact"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' +-+-+-+-+-+-+-+-+-+-+-+-
' +-+- Autor: MANOLO  +-+-
' +-+-+-+-+-+-+-+-+-+-+-+-

Option Explicit

Public DatosADevolverBusqueda As String    'Tindrà el nº de text que vol que torne, empipat
Public Event DatoSeleccionado(CadenaSeleccion As String)


Public numfactu As Long
Public LetraSerie As String
Public Tipo As Byte ' 0 schfac normal
                    ' 1 schfacr ajena para el Regaixo
                    ' 2 schfac1 historico de facturas 1

Private Modo As Byte
'*************** MODOS ********************
'   0.-  Formulari net sense cap camp ple
'   1.-  Preparant per a fer la búsqueda
'   2.-  Ja tenim registres i els anem a recorrer
'        ,podem editar-los Edició del camp
'   3.-  Inserció de nou registre
'   4.-  Modificar
'   5.-  Manteniment Llinies

'***Variables comuns a tots els formularis*****

Dim ModoLineas As Byte
'1.- Añadir,  2.- Modificar,  3.- Borrar,  0.-Pasar control a Lineas

Dim NumTabMto As Integer 'Indica quin nº de Tab està en modo Mantenimient
Dim TituloLinea As String 'Descripció de la llínia que està en Mantenimient
Dim PrimeraVez As Boolean

Private CadenaConsulta As String 'SQL de la taula principal del formulari
Private Ordenacion As String
Private NomTabla As String  'Nom de la taula

Private kCampo As Integer
'-------------------------------------------------------------------------
Private HaDevueltoDatos As Boolean

Dim btnPrimero As Byte 'Variable que indica el nº del Botó PrimerRegistro en la Toolbar1
'Dim CadAncho() As Boolean  'array, per a quan cridem al form de llínies
Dim indice As Byte 'Indice del text1 donde se ponen los datos devueltos desde otros Formularios de Mtos

Private WithEvents frmC As frmCal 'calendario fechas
Attribute frmC.VB_VarHelpID = -1
Private WithEvents frmB As frmBuscaGrid
Attribute frmB.VB_VarHelpID = -1

Private WithEvents frmcli As frmManClien
Attribute frmcli.VB_VarHelpID = -1
Private WithEvents frmCol As frmManCoope
Attribute frmCol.VB_VarHelpID = -1
Private WithEvents frmArt As frmManArtic
Attribute frmArt.VB_VarHelpID = -1
Private WithEvents frmFPa As frmManFpago
Attribute frmFPa.VB_VarHelpID = -1
Private WithEvents frmDep As frmManDpto
Attribute frmDep.VB_VarHelpID = -1
Private WithEvents frmTipIVA As frmTipIVAConta  'Tipos de IVA de la contabilidad
Attribute frmTipIVA.VB_VarHelpID = -1
Private WithEvents frmList As frmListado
Attribute frmList.VB_VarHelpID = -1

Dim ClienteAnt As String
Dim FormaPagoAnt As String
Dim ModoModificar As Boolean
Dim ModificaImportes As Boolean ' variable que me indica q hay que modificar lineas de la factura de contabilidad
                                ' y cobros en la tesoreria

'GENERALES PARA PASARLE A CRYSTAL REPORT
Private cadFormula As String 'Cadena con la FormulaSelection para Crystal Report
Private cadParam As String 'Cadena con los parametros para Crystal Report
Private numParam As Byte 'Numero de parametros que se pasan a Crystal Report
Private cadSelect As String 'Cadena para comprobar si hay datos antes de abrir Informe
Private cadTitulo As String 'Titulo para la ventana frmImprimir
Private cadNombreRPT As String 'Nombre del informe

Dim TipForpa As String
Dim TipForpaAnt As String

' utilizado para buscar por checks
Private BuscaChekc As String

Dim CadenaBorrado As String

Dim OpcionListado As Byte

Private Sub btnBuscar_Click(Index As Integer)
    ' els formularis als que crida son d'una atra BDA
    TerminaBloquear
    
    Select Case Index
        Case 0 'Artículos
            Set frmArt = New frmManArtic
            frmArt.DatosADevolverBusqueda = "0|1|"
            frmArt.CodigoActual = txtAux(9).Text
            frmArt.Show vbModal
            Set frmArt = Nothing
            
    End Select
    
    PonerFoco txtAux(9)
    If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Data1, 1
End Sub

Private Sub Check1_Click(Index As Integer)
    If Modo = 1 Then
        'Buscqueda
        If InStr(1, BuscaChekc, "check1(" & Index & ")") = 0 Then BuscaChekc = BuscaChekc & "check1(" & Index & ")|"
    End If
End Sub

Private Sub Check1_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

'Private Sub Check1_Click()
'    If Modo = 1 Then
'        'Buscqueda
'        If InStr(1, BuscaChekc, "check1") = 0 Then BuscaChekc = BuscaChekc & "check1|"
'    End If
'End Sub

'Private Sub Check1_KeyPress(KeyAscii As Integer)
'    KEYpress KeyAscii
'End Sub

Private Sub cmdAceptar_Click()
Dim b As Boolean
Dim vTabla As String
Dim CtaClie As String

    Screen.MousePointer = vbHourglass
    On Error GoTo Error1
    
    ModoModificar = False
    b = True
    Select Case Modo
        Case 1  'BUSQUEDA
            HacerBusqueda
            
'        Case 3 'INSERTAR
'            If DatosOk Then
'                If InsertarDesdeForm2(Me, 1) Then
'                    Data1.RecordSource = "Select * from " & NomTabla & Ordenacion
'                    PosicionarData
'                End If
'            Else
'                ModoLineas = 0
'            End If
'
        Case 4  'MODIFICAR
            If Not DatosOk Then
                ModoLineas = 0
                Screen.MousePointer = vbDefault
                Exit Sub
            Else
                ModoModificar = True
                
                Conn.BeginTrans
                If vParamAplic.NumeroConta <> 0 Then ConnConta.BeginTrans
                
                If CadenaBorrado <> "" Then
                    Conn.Execute CadenaBorrado
                    CadenaBorrado = ""
                    EliminarLinea
                End If
                
                
                If ModificaDesdeFormulario2(Me, 1) Then
                    If vParamAplic.NumeroConta <> 0 And Check1(1).Value = 1 Then
                        'solo en el caso de que este contabilizada
                        If Val(ClienteAnt) <> Val(Text1(3).Text) Then
                            CtaClie = ""
                            CtaClie = DevuelveDesdeBDNew(cPTours, "ssocio", "codmacta", "codsocio", Text1(3).Text, "N")
                            b = ModificaClienteFacturaContabilidad(Text1(0).Text, Text1(1).Text, Text1(2).Text, CtaClie, Tipo)
                        End If
' 09022007 ya no dejo modificar la forma de pago
'                        If Val(FormaPagoAnt) <> Val(Text1(5).Text) Then _
'                            ModificaFormaPagoTesoreria Text1(0).Text, Text1(1).Text, Text1(2).Text, Text1(5).Text, FormaPagoAnt, TipForpa, TipForpaAnt
                            
                        If ModificaImportes And b Then
                            BorrarTMPErrFact
                            '[Monica]24/07/2013
                            Select Case Tipo
                                Case 0:
                                    vTabla = "schfac"
                                Case 1:
                                    vTabla = "schfacr"
                                Case 2:
                                    vTabla = "schfac1"
                            End Select
                            b = ModificaImportesFacturaContabilidad(Text1(0).Text, Text1(1).Text, Text1(2).Text, Text1(18).Text, Text1(5).Text, vTabla)
                            ModificaImportes = False
                        End If
                    End If
                    TerminaBloquear
                    PosicionarData "letraser = '" & Text1(0).Text & "' and numfactu = " & Text1(1).Text & " and fecfactu = " & DBSet(Text1(2).Text, "F")
                End If
            End If
            
        Case 5 'LLINIES
            Select Case ModoLineas
'                Case 1 'afegir llinia
'                    InsertarLinea
                Case 2 'modificar llinies
                    ModificarLinea
                    PosicionarData "letraser = '" & Text1(0).Text & "' and numfactu = " & Text1(1).Text & " and fecfactu = " & DBSet(Text1(2).Text, "F")
                    Screen.MousePointer = vbDefault
                    Exit Sub
                Case 3 'eliminar llinies
                    ModificarLinea
                    PosicionarData "letraser = '" & Text1(0).Text & "' and numfactu = " & Text1(1).Text & " and fecfactu = " & DBSet(Text1(2).Text, "F")
            
            End Select
    End Select
    Screen.MousePointer = vbDefault
 
Error1:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Or Not b Then
        MsgBox Err.Number & ": " & Err.Description, vbExclamation
        If ModoModificar Then
            Conn.RollbackTrans
            ConnConta.RollbackTrans
        End If
    Else
        If ModoModificar Then
            Conn.CommitTrans
            ConnConta.CommitTrans
        End If
    End If
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
Dim sql2 As String

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
        .Buttons(10).Image = 16 ' Rectificativas
        'el 10 i el 11 son separadors
        .Buttons(12).Image = 10  'Imprimir
        .Buttons(13).Image = 11  'Salir
        'el 14 i el 15 son separadors
        .Buttons(btnPrimero).Image = 6  'Primero
        .Buttons(btnPrimero + 1).Image = 7 'Anterior
        .Buttons(btnPrimero + 2).Image = 8 'Siguiente
        .Buttons(btnPrimero + 3).Image = 9 'Último
    End With
    
    'ICONITOS DE LAS BARRAS EN LOS TABS DE LINEA
    For I = 0 To ToolAux.Count - 1
        With Me.ToolAux(I)
            '.ImageList = frmPpal.imgListComun_VELL
            '  ### [Monica] 02/10/2006 acabo de comentarlo
            '.HotImageList = frmPpal.imgListComun_OM16
            '.DisabledImageList = frmPpal.imgListComun_BN16
            .ImageList = frmPpal.imgListComun16
            .Buttons(1).Image = 3   'Insertar
            .Buttons(2).Image = 4   'Modificar
            .Buttons(3).Image = 5   'Borrar
        End With
    Next I
    
    'cargar IMAGES de busqueda
    For I = 0 To Me.imgBuscar.Count - 1
        Me.imgBuscar(I).Picture = frmPpal.imgListImages16.ListImages(1).Picture
    Next I
   
    
    LimpiarCampos   'Limpia los campos TextBox
    For I = 0 To DataGridAux.Count - 1 'neteje tots els grids de llinies
        DataGridAux(I).ClearFields
    Next I
    
    '## A mano
    Select Case Tipo
        Case 0:
            NomTabla = "schfac"
        Case 1:
            NomTabla = "schfacr"
        
            Me.Caption = Me.Caption & " Ajenas"
            
        Case 2:
            NomTabla = "schfac1"
        
            Me.Caption = Me.Caption & " 1"
            
    End Select
    CambiarTags Tipo
    Ordenacion = " ORDER BY letraser, numfactu, fecfactu "
    
    'Vemos como esta guardado el valor del check
    chkVistaPrevia(0).Value = CheckValueLeer(Name)
    
    '[Monica]15/02/2011: El impuesto sigaus solo lo voy a mostrar en el caso de Pobla del Duc
    Label1(0).visible = (vParamAplic.Cooperativa = 4)
    Text1(20).visible = (vParamAplic.Cooperativa = 4)
    '[Monica]15/02/2011: fin
    
    
    
    'ASignamos un SQL al DATA1
    Data1.ConnectionString = Conn
    sql2 = "Select * from " & NomTabla & " where numfactu is null "
    Data1.RecordSource = sql2
    Data1.Refresh
        
    If DatosADevolverBusqueda = "" Then
        PonerModo 0
    Else
        PonerModo 1
        Text1(0).BackColor = vbYellow 'letraser
    End If
    
    ModoLineas = 0
    
    For I = 0 To DataGridAux.Count - 1
        CargaGrid I, (Modo = 2) 'carregue els datagrids de llinies
    Next I
    
    If LetraSerie <> "" Then
        Text1(0).Text = LetraSerie
        Text1(1).Text = numfactu
        PonerModo 1
        cmdAceptar_Click
    End If

    Me.SSTab1.Tab = 0


End Sub

Private Sub LimpiarCampos()
    On Error Resume Next

    limpiar Me   'Metodo general: Limpia los controles TextBox
    lblIndicador.Caption = ""
    Me.Check1(0).Value = 0
    Me.Check1(1).Value = 0
    'Aqui va el especifico de cada form es
    '### a mano
    If Err.Number <> 0 Then Err.Clear
End Sub

'----------------------------------------------------------------
'----------------------------------------------------------------
'   En PONERMODO se habilitan, o no, los diverso campos del
'   formulario en funcion del modo en k vayamos a trabajar
Private Sub PonerModo(Kmodo As Byte)
Dim I As Integer, Numreg As Byte
Dim b As Boolean
On Error GoTo EPonerModo
 
    Modo = Kmodo
    BuscaChekc = ""
    
    PonerIndicador lblIndicador, Modo, ModoLineas
       
    'Modo 2. Hay datos y estamos visualizandolos
    '=========================================
    'Ponemos visible, si es formulario de busqueda, el boton regresar cuando hay datos
    If DatosADevolverBusqueda <> "" Then
        cmdRegresar.visible = (Modo = 2)
    Else
        cmdRegresar.visible = False
    End If
    
    '=======================================
    b = (Modo = 2)
    'Poner Flechas de desplazamiento visibles
    Numreg = 1
    If Not Data1.Recordset.EOF Then
        If Data1.Recordset.RecordCount > 1 Then Numreg = 2 'Solo es para saber q hay + de 1 registro
    End If
    DesplazamientoVisible Me.Toolbar1, btnPrimero, b, Numreg
    

    b = Modo <> 0 And Modo <> 2
    cmdCancelar.visible = b
    cmdAceptar.visible = b
    
    '---------------------------------------------
    
    'Bloquea los campos Text1 si no estamos modificando/Insertando Datos
    'Si estamos en Insertar además limpia los campos Text1
    BloquearText1 Me, Modo
    BloquearChecks Me, Modo
    
    BloquearImgBuscar Me, Modo, ModoLineas
       
    'Bloquear los campos de clave primaria, NO se puede modificar
    b = (Modo = 3) Or (Modo = 1)  'solo al insertar/buscar estará activo
    For I = 0 To 2
        BloquearTxt Text1(I), Not b
    Next I
    'Los % de IVA siempre bloqueados
    BloquearTxt Text1(8), True
    BloquearTxt Text1(12), True
    BloquearTxt Text1(16), True
    'El total de la factura siempre bloqueado
    BloquearTxt Text1(18), True
    BloquearTxt Text1(19), True
    BloquearTxt Text1(20), True
    
    '09/02/2007 no dejo modificar la forma de pago
    BloquearTxt Text1(5), Not b
    
    
    Text1(18).BackColor = &HCAE3FD
    Text1(19).BackColor = &HC0C0FF
    Text1(20).BackColor = &HC0C0FF

    b = (Modo = 3) Or (Modo = 1) Or (Modo = 4)
    Me.imgBuscar(0).Enabled = b
    Me.imgBuscar(0).visible = b
    Me.imgBuscar(1).Enabled = b
    Me.imgBuscar(1).visible = b
    Me.imgBuscar(2).Enabled = ((Modo = 3) Or (Modo = 1))
    Me.imgBuscar(2).visible = ((Modo = 3) Or (Modo = 1))
    
    
    'Imagen Calendario fechas
    b = (Modo = 3 Or Modo = 4 Or Modo = 1 Or (Modo = 5 And (ModoLineas = 1 Or ModoLineas = 2)))
    Me.imgFec(2).Enabled = (Modo = 3 Or Modo = 1) 'es clave, solo al insertar o buscar
    Me.imgFec(2).visible = (Modo = 3 Or Modo = 1) 'es clave, solo al insertar o buscar
            
    chkVistaPrevia(0).Enabled = (Modo <= 2)
    
    'Poner el tamaño de los campos. Si es modo Busqueda el MaxLength del campo
    'debe ser mayor.
    PonerLongCampos
                          
    If (Modo < 2) Or (Modo = 3) Then
        For I = 0 To DataGridAux.Count - 1
            CargaGrid I, False
        Next I
    End If
    
    b = (Modo = 4) Or (Modo = 2)
    For I = 0 To DataGridAux.Count - 1
        DataGridAux(I).Enabled = b
    Next I
    
    b = (Modo = 4)
    FrameTotFactu.Enabled = Not b
    
    b = (Modo = 5)
    
    For I = 21 To 24
        BloquearTxt Text1(I), (Modo <> 1)
    Next I
    
    PonerModoOpcionesMenu (Modo) 'Activar opciones menu según modo
    PonerOpcionesMenu   'Activar opciones de menu según nivel
                        'de permisos del usuario

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
'Activas unas Opciones de Menu y Toolbar según el nivel de usuario
    PonerOpcionesMenuGeneral Me
    PonerOpcionesMenuGeneralNew Me
End Sub

Private Sub PonerModoOpcionesMenu(Modo)
'Activas unas Opciones de Menu y Toolbar según el modo en que estemos
Dim b As Boolean, bAux As Boolean
Dim I As Byte

    '-----  TOOLBAR DE LA CABECERA
    '------------------------------------------
    'b = (Modo = 2 Or Modo = 0 Or Modo = 1)
    b = (Modo = 2 Or Modo = 0)
    'Buscar
    Toolbar1.Buttons(3).Enabled = b
    Me.mnBuscar.Enabled = b
    'Ver Todos
    Toolbar1.Buttons(4).Enabled = b
    Me.mnVerTodos.Enabled = b
    
    'Insertar
    Toolbar1.Buttons(7).Enabled = b
    Me.mnNuevo.Enabled = b
    
    
    b = (Modo = 2 And Data1.Recordset.RecordCount > 0)
    'Modificar
    Toolbar1.Buttons(8).Enabled = b
    Me.mnModificar.Enabled = b
    'eliminar
    Toolbar1.Buttons(9).Enabled = b
    Me.mnEliminar.Enabled = b
    'rectificativas
    Toolbar1.Buttons(10).Enabled = b And (Tipo = 0) And EsFacturaRectificable(Text1(0).Text)
    Me.mnRectificar.Enabled = b And (Tipo = 0) And EsFacturaRectificable(Text1(0).Text)
   
    'Imprimir
    'VRS:2.0.1(3)
    Toolbar1.Buttons(12).Enabled = (Modo = 2)
    Me.mnImprimir.Enabled = (Modo = 2)
    '-----------  LINEAS
    ' *** MEU: botons de les llínies de cuentas bancarias,
    ' només es poden gastar quan inserte o modifique clients ****
    'b = (Modo = 3 Or Modo = 4)
    b = (Modo = 3 Or Modo = 4 Or Modo = 2)
    For I = 0 To ToolAux.Count - 1
        ToolAux(I).Buttons(1).Enabled = b
        If b Then bAux = (b And Me.AdoAux(I).Recordset.RecordCount > 0)
        ToolAux(I).Buttons(2).Enabled = bAux
        ToolAux(I).Buttons(3).Enabled = bAux
    Next I
    'Imprimir en pestaña Comisiones de Productos
'    ToolAux(2).Buttons(6).Enabled = (Modo = 2) Or (Modo = 3) Or (Modo = 4) Or (Modo = 5 And ModoLineas = 0)
    ' ************************************************************
End Sub

Private Sub Desplazamiento(Index As Integer)
'Botones de Desplazamiento de la Toolbar
'Para desplazarse por los registros de control Data
    If Data1.Recordset.EOF Then Exit Sub
    DesplazamientoData Data1, Index
    PonerCampos
End Sub

Private Function MontaSQLCarga(Index As Integer, enlaza As Boolean) As String
'--------------------------------------------------------------------
' MontaSQlCarga:
'   Basándose en la información proporcionada por el vector de campos
'   crea un SQl para ejecutar una consulta sobre la base de datos que los
'   devuelva.
' Si ENLAZA -> Enlaza con el data1
'           -> Si no lo cargamos sin enlazar a ningun campo
'--------------------------------------------------------------------
Dim SQL As String
Dim tabla As String
    
    Select Case Index
        Case 0 'Lineas de factura
            Select Case Tipo
                Case 0:
                    tabla = "slhfac"
                    SQL = "SELECT letraser,numfactu,fecfactu,numlinea,numalbar,fecalbar,"
                    SQL = SQL & "horalbar, codturno, numtarje, slhfac.codartic, sartic.nomartic, "
                    SQL = SQL & "cantidad, preciove, implinea "
                    SQL = SQL & " FROM slhfac, sartic "
                    SQL = SQL & " WHERE slhfac.codartic = sartic.codartic "
        
                    If enlaza Then
                        SQL = SQL & " AND " & ObtenerWhereCab(False)
                    Else
                        SQL = SQL & " AND numfactu is null "
                    End If
                    SQL = SQL & " ORDER BY " & tabla & ".numlinea "
                Case 1:
                    ' facturacion ajena
                    tabla = "slhfacr"
                    SQL = "SELECT letraser,numfactu,fecfactu,numlinea,numalbar,fecalbar,"
                    SQL = SQL & "horalbar, codturno, numtarje, slhfacr.codartic, sartic.nomartic, "
                    SQL = SQL & "cantidad, preciove, implinea "
                    SQL = SQL & " FROM slhfacr, sartic "
                    SQL = SQL & " WHERE slhfacr.codartic = sartic.codartic "
        
                    If enlaza Then
                        SQL = SQL & " AND " & ObtenerWhereCab(False)
                    Else
                        SQL = SQL & " AND numfactu is null "
                    End If
                    SQL = SQL & " ORDER BY " & tabla & ".numlinea "
            
                Case 2:
                    ' historico de facturas
                    tabla = "slhfac1"
                    SQL = "SELECT letraser,numfactu,fecfactu,numlinea,numalbar,fecalbar,"
                    SQL = SQL & "horalbar, codturno, numtarje, slhfac1.codartic, sartic.nomartic, "
                    SQL = SQL & "cantidad, preciove, implinea "
                    SQL = SQL & " FROM slhfac1, sartic "
                    SQL = SQL & " WHERE slhfac1.codartic = sartic.codartic "
        
                    If enlaza Then
                        SQL = SQL & " AND " & ObtenerWhereCab(False)
                    Else
                        SQL = SQL & " AND numfactu is null "
                    End If
                    SQL = SQL & " ORDER BY " & tabla & ".numlinea "
            End Select
    End Select
    MontaSQLCarga = SQL
End Function

Private Sub frmB_Selecionado(CadenaDevuelta As String)
    Dim CadB As String
    Dim Aux As String
    
    If CadenaDevuelta <> "" Then
        HaDevueltoDatos = True
        Screen.MousePointer = vbHourglass
        'Sabemos que campos son los que nos devuelve
        'Creamos una cadena consulta y ponemos los datos
        CadB = ""
        Aux = ValorDevueltoFormGrid(Text1(0), CadenaDevuelta, 1) 'numfactu
        CadB = Aux
        Aux = ValorDevueltoFormGrid(Text1(1), CadenaDevuelta, 2) 'fecfactu
        CadB = CadB & " AND " & Aux
        Aux = ValorDevueltoFormGrid(Text1(2), CadenaDevuelta, 3) 'codprove
        CadB = CadB & " AND " & Aux
        '   Como la clave principal es unica, con poner el sql apuntando
        '   al valor devuelto sobre la clave ppal es suficiente
        CadenaConsulta = "select * from " & NomTabla & " WHERE " & CadB & " " & Ordenacion
        PonerCadenaBusqueda
        Screen.MousePointer = vbDefault
    End If
End Sub

Private Sub frmC_Selec(vFecha As Date)
    'Fecha
    Text1(CByte(imgFec(2).Tag)).Text = Format(vFecha, "dd/mm/yyyy")
End Sub

Private Sub frmArt_DatoSeleccionado(CadenaSeleccion As String)
'Articulos
    txtAux(9).Text = RecuperaValor(CadenaSeleccion, 1) 'codartic
    txtAux2(0).Text = RecuperaValor(CadenaSeleccion, 2) 'nomartic
End Sub

Private Sub frmDep_DatoSeleccionado(CadenaSeleccion As String)
'Mantenimiento de Departamentos
    Text1(26).Text = RecuperaValor(CadenaSeleccion, 1) 'coddepar
    FormateaCampo Text1(26)
    text2(26).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmFpa_DatoSeleccionado(CadenaSeleccion As String)
'Mantenimiento de Formas de Pago
    Text1(5).Text = RecuperaValor(CadenaSeleccion, 1) 'codforpa
    FormateaCampo Text1(5)
    text2(5).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmCli_DatoSeleccionado(CadenaSeleccion As String)
'Mantenimiento de Clientes
Dim Cad As String, Datos As String

    Text1(3).Text = RecuperaValor(CadenaSeleccion, 1) 'codclien
    FormateaCampo Text1(3)
    text2(3).Text = RecuperaValor(CadenaSeleccion, 2) 'nomclien
     
    'recuperar el Colectivo del cliente
    Datos = DevuelveDesdeBDNew(cPTours, "ssocio", "codcoope", "codsocio", Text1(3).Text, "N")
    If Datos <> "" Then
        Text1(4).Text = Datos
        FormateaCampo Text1(4)
        Text1_LostFocus (4)
    Else
        Text1(4).Text = ""
        text2(4).Text = ""
    End If
End Sub

Private Sub frmList_DatoSeleccionado(CadenaSeleccion As String)

End Sub

Private Sub frmList_RectificarFactura(Cliente As String, Observaciones As String)
'    If CrearFacturaRectificativa(Text1(0).Text, Text1(1).Text, Text1(2).Text, observaciones, cliente) = 0 Then
'        MsgBox "Proceso realizado correctamente", vbExclamation
'    End If
End Sub

Private Sub frmTipIVA_DatoSeleccionado(CadenaSeleccion As String)
'Tipos de IVA (de la Contabilidad)
    Text1(indice).Text = RecuperaValor(CadenaSeleccion, 1) 'codigiva
    FormateaCampo Text1(indice)
    Text1(indice + 3).Text = RecuperaValor(CadenaSeleccion, 3) '% iva
End Sub

Private Sub imgBuscar_Click(Index As Integer)
   'Screen.MousePointer = vbHourglass
    TerminaBloquear
    
    Select Case Index

        Case 0 'Cliente
            indice = 3
            Set frmcli = New frmManClien
            frmcli.DatosADevolverBusqueda = "0|1|"
            frmcli.CodigoActual = Text1(3).Text
            frmcli.Show vbModal
            Set frmcli = Nothing
            
        Case 1 'Colectivo
            indice = 4
            Set frmCol = New frmManCoope
            frmCol.DatosADevolverBusqueda = "0|1|"
            frmCol.CodigoActual = Text1(4).Text
            frmCol.Show vbModal
            Set frmCol = Nothing
            
        Case 2 'forma de pago
            indice = 5
            Set frmFPa = New frmManFpago
            frmFPa.DatosADevolverBusqueda = "0|1|"
            frmFPa.DeConsulta = True
            frmFPa.CodigoActual = Text1(5).Text
            frmFPa.Show vbModal
            Set frmFPa = Nothing
            
        Case 3, 4, 5 'tiposd de IVA (de la contabilidad)
            If Index = 3 Then Let indice = 7
            If Index = 4 Then Let indice = 11
            If Index = 5 Then Let indice = 15
            Set frmTipIVA = New frmTipIVAConta
            frmTipIVA.DatosADevolverBusqueda = "0|1|2|"
            frmTipIVA.CodigoActual = Text1(indice).Text
            frmTipIVA.Show vbModal
            Set frmTipIVA = Nothing
            
            
        Case 6 ' departamento
            indice = 6
            Set frmDep = New frmManDpto
            frmDep.DatosADevolverBusqueda = "0|1|"
            frmDep.CodigoActual = Text1(5).Text
            frmDep.Show vbModal
            Set frmDep = Nothing
        
            
    End Select
    
'    If Index = 3 Or Index = 4 Then
'        PonerFoco txtAux(Indice)
'    Else
        PonerFoco Text1(indice)
'    End If
    If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Data1, 1
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

    ' es desplega baix i cap a la dreta
    'frmC.Left = esq + imgFec(Index).Parent.Left + 30
    'frmC.Top = dalt + imgFec(Index).Parent.Top + imgFec(Index).Height + menu - 40
    
    ' es desplega dalt i cap a la esquerra
    frmC.Left = esq + imgFec(Index).Parent.Left - frmC.Width + imgFec(Index).Width + 40
    frmC.Top = dalt + imgFec(Index).Parent.Top - frmC.Height + menu - 25
       
    ' ***canviar l'index de imgFec pel 1r index de les imagens de buscar data***
    imgFec(2).Tag = Index 'independentment de les dates que tinga, sempre pose l'index en la 27
    If Text1(Index).Text <> "" Then frmC.NovaData = Text1(Index).Text

    frmC.Show vbModal
    Set frmC = Nothing
    PonerFoco Text1(CByte(imgFec(2).Tag))
    ' ***************************
End Sub

Private Sub mnBuscar_Click()
    BotonBuscar
    Me.Check1(0).Value = 0
    Me.Check1(1).Value = 0
End Sub

Private Sub mnEliminar_Click()
    BotonEliminar
End Sub

Private Sub mnImprimir_Click()
Dim Letra As String

    'VRS:2.0.1(3): añadido el boton de imprimir
    cadTitulo = "Reimpresion de Facturas"

    ' ### [Monica] 11/09/2006
    '****************************
    Dim indRPT As Byte 'Indica el tipo de Documento en la tabla "scryst"
    Dim nomDocu As String 'Nombre de Informe rpt de crystal

    indRPT = 1 'Facturas Clientes

    If Not PonerParamRPT(indRPT, cadParam, numParam, nomDocu) Then Exit Sub
    'Nombre fichero .rpt a Imprimir
    If Tipo = 1 Then
'02/03/2007 he duplicado documentos
'        nomDocu = Replace(nomDocu, ".rpt", "Ajena.rpt")
        
        nomDocu = Replace(nomDocu, ".rpt", "Aj" & "C" & Format(Data1.Recordset!codcoope, "00") & ".rpt")

        cadTitulo = cadTitulo & " Ajenas"
    End If
    '30/11/2010: vemos si es una factura cepsa
    If vParamAplic.Cooperativa = 4 Then
        Letra = DevuelveValor("select letraser from stipom where codtipom = 'FAC'")
        If Letra = Trim(Text1(0).Text) Then
            nomDocu = Replace(nomDocu, ".rpt", "Cepsa.rpt")
            cadTitulo = cadTitulo & " de Cepsa"
        End If
    End If
    
    
    frmImprimir.NombreRPT = nomDocu
    ' he añadido estas dos lineas para que llame al rpt correspondiente

    cadNombreRPT = nomDocu  ' "rFactgas.rpt"
    
    cadFormula = "({" & NomTabla & ".letraser} = """ & Text1(0).Text & """) AND ({" & NomTabla & ".numfactu} = " & Text1(1).Text & ") and ({" & NomTabla & ".fecfactu} = cdate(""" & Text1(2).Text & """)) "
    
    '23022007 Monica: la separacion de la bonificacion solo la quieren en Alzira
    '[Monica]09/01/2013: Nueva cooperativa de Ribarroja
    If vParamAplic.Cooperativa = 1 Or vParamAplic.Cooperativa = 5 Then
        cadFormula = cadFormula & " and {slhfac.numalbar} <> 'BONIFICA'" ' AND ({ssocio.impfactu}<=1)"
        OpcionListado = 1
    Else
        OpcionListado = 0
    End If
    
    cadParam = "|pEmpresa=" & vEmpresa.nomEmpre '& "|pCodigoISO="11112"|pCodigoRev="01"|
    
    LlamarImprimir
End Sub

Private Sub mnModificar_Click()

    'Comprobaciones
    '--------------
    If Data1.Recordset.EOF Then Exit Sub
    If Data1.Recordset.RecordCount < 1 Then Exit Sub
    
    'El registre de codi 0 no es pot Modificar ni Eliminar
    ' ### [Monica] 27/09/2006
    ' quitamos el control de no poder modificar ni eliminar si es 0
    'If EsCodigoCero(CStr(Data1.Recordset.Fields(0).Value), FormatoCampo(text1(0))) Then Exit Sub
    
    ' ### [Monica] 27/09/2006
    ' solo podemos modificar en el caso de que haya contabilidad si la factura es modificable
    If vParamAplic.NumeroConta <> 0 And Not FacturaModificable(Text1(0).Text, Text1(1).Text, Text1(2).Text, Check1(1).Value) Then Exit Sub
    
    'Preparar para modificar
    '-----------------------
    If BLOQUEADesdeFormulario2(Me, Data1, 1) Then BotonModificar
End Sub

'Private Sub mnNuevo_Click()
'     BotonAnyadir
'End Sub

Private Sub mnRectificar_Click()

    'Comprobaciones
    '--------------
    If Data1.Recordset.EOF Then Exit Sub
    If Data1.Recordset.RecordCount < 1 Then Exit Sub
    
    'El registre de codi 0 no es pot Modificar ni Eliminar
    ' ### [Monica] 27/09/2006
    ' quitamos el control de no poder modificar ni eliminar si es 0
    'If EsCodigoCero(CStr(Data1.Recordset.Fields(0).Value), FormatoCampo(text1(0))) Then Exit Sub
    
    ' ### [Monica] 27/09/2006
    ' solo podemos modificar en el caso de que haya contabilidad si la factura es modificable
    If vParamAplic.NumeroConta <> 0 And Not FacturaModificable(Text1(0).Text, Text1(1).Text, Text1(2).Text, Check1(1).Value) Then Exit Sub
    
    'Preparar para modificar
    '-----------------------
    If Not FacturaModificable(Text1(0).Text, Text1(1).Text, Text1(2).Text, Check1(1).Value) Then Exit Sub
    If BLOQUEADesdeFormulario2(Me, Data1, 1) Then BotonRectificar
End Sub



Private Sub mnSalir_Click()
    Unload Me
End Sub

Private Sub mnVerTodos_Click()
    BotonVerTodos
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 3  'Buscar
           mnBuscar_Click
        Case 4  'Todos
            mnVerTodos_Click
'        Case 7  'Nuevo
'            mnNuevo_Click
        Case 8  'Modificar
            mnModificar_Click
'        Case 9  'Borrar
'            mnEliminar_Click
        Case 10 'Rectificativa
            mnRectificar_Click
        Case 12 'Imprimir
            mnImprimir_Click
        Case 13    'Salir
            mnSalir_Click
        Case btnPrimero To btnPrimero + 3 'Flechas Desplazamiento
            Desplazamiento (Button.Index - btnPrimero)
    End Select
End Sub

Private Sub BotonBuscar()
    'Buscar
    If Modo <> 1 Then
        LimpiarCampos
        'Poner los grid sin apuntar a nada
        'LimpiarDataGrids
        PonerModo 1
        'Si pasamos el control aqui lo ponemos en amarillo
        PonerFoco Text1(0)
        Text1(0).BackColor = vbYellow
    Else
        HacerBusqueda
        If Data1.Recordset.EOF Then
            Text1(kCampo).Text = ""
            Text1(kCampo).BackColor = vbYellow
            PonerFoco Text1(kCampo)
        End If
    End If
End Sub

Private Sub HacerBusqueda()
Dim CadB As String

    CadB = ObtenerBusqueda2(Me, BuscaChekc)
    '++monica: 07/04/08 el busca check no fucntionaba en las ajenas de regaixo
    Select Case Tipo
        Case 1:
            CadB = Replace(CadB, "schfac.", "schfacr.")
        Case 2:
            CadB = Replace(CadB, "schfac.", "schfac1.")
    End Select
    '--monica
    
    If chkVistaPrevia(0) = 1 Then
        MandaBusquedaPrevia CadB
    ElseIf CadB <> "" Then
        'Se muestran en el mismo form
        CadenaConsulta = "select * from " & NomTabla & " WHERE " & CadB & " " & Ordenacion
        PonerCadenaBusqueda
    Else
        PonerFoco Text1(0)
    End If
End Sub

Private Sub MandaBusquedaPrevia(CadB As String)
Dim Cad As String
        'Llamamos a al form
        '##A mano
        Cad = ""
        Cad = Cad & ParaGrid(Text1(0), 7, "Serie")
        Cad = Cad & ParaGrid(Text1(1), 16, "Nº Fact.")
        Cad = Cad & ParaGrid(Text1(2), 15, "Fecha")
'        cad = cad & ParaGrid(Text1(3), 35, "Proveedor")
        'los ponemos a mano:
        Cad = Cad & "Cliente.|" & NomTabla & ".codsocio|N|" & FormatoCampo(Text1(3)) & "|12·"
        Cad = Cad & "Nom. Cliente|nomsocio|T||50·"
'        cad = cad & "Empr.|factproc.codempre|N|000|8·"
        If Cad <> "" Then
            Screen.MousePointer = vbHourglass
            Set frmB = New frmBuscaGrid
            frmB.vCampos = Cad
            frmB.vTabla = NomTabla & " INNER JOIN ssocio ON " & NomTabla & ".codsocio=ssocio.codsocio "
            frmB.vSQL = CadB
            HaDevueltoDatos = False
            '###A mano
            frmB.vDevuelve = "0|1|2|4|"
            frmB.vTitulo = "Facturas Clientes"
            frmB.vSelElem = 0
            '#
            frmB.Show vbModal
            Set frmB = Nothing
            'Si ha puesto valores y tenemos que es formulario de busqueda entonces
            'tendremos que cerrar el form lanzando el evento
            If HaDevueltoDatos Then
                If (Not Data1.Recordset.EOF) And DatosADevolverBusqueda <> "" Then _
                    cmdRegresar_Click
            Else   'de ha devuelto datos, es decir NO ha devuelto datos
                PonerFoco Text1(kCampo)
            End If
        End If
End Sub

Private Sub cmdRegresar_Click()
Dim Cad As String
Dim Aux As String
Dim I As Integer
Dim J As Integer

    If Data1.Recordset.EOF Then
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
            Cad = Cad & Text1(J).Text & "|"
        End If
    Loop Until I = 0
    RaiseEvent DatoSeleccionado(Cad)
    Unload Me
End Sub

Private Sub PonerCadenaBusqueda()
    Screen.MousePointer = vbHourglass
    On Error GoTo EEPonerBusq
    
    Data1.RecordSource = CadenaConsulta
    Data1.Refresh
    If Data1.Recordset.RecordCount <= 0 Then
        MsgBox "No hay ningún registro en la tabla " & NomTabla, vbInformation
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
'Ver todos
Dim I As Integer

    LimpiarCampos 'Limpia los Text1
    
    For I = 0 To DataGridAux.Count - 1 'Limpias los DataGrid
        CargaGrid I, False
    Next I
    
    If chkVistaPrevia(0).Value = 1 Then
        MandaBusquedaPrevia ""
    Else
        CadenaConsulta = "Select * from " & NomTabla & Ordenacion
        PonerCadenaBusqueda
    End If
End Sub

'Private Sub BotonAnyadir()
''Añadir registro en tabla de expedientes individuales: expincab (Cabecera)
'
'    LimpiarCampos 'Vacía los TextBox
'    'Poner los grid sin apuntar a nada
''    LimpiarDataGrids
'
'    'Añadiremos el boton de aceptar y demas objetos para insertar
'    PonerModo 3
'
'    'Quan afegixc pose en Fecha
'    Text1(2).Text = Format(Now, "dd/mm/yyyy")
'
'    'Total Factura (por defecto=0)
'    Text1(18).Text = "0"
'    Text1(19).Text = "0"
'
'    'em posicione en el 1r tab
'    PonerFoco Text1(0)
'End Sub

Private Sub BotonModificar()
    
    'Añadiremos el boton de aceptar y demas objetos para insertar
    PonerModo 4
    
    ' ### [Monica] 27/09/2006
    ' me guardo los valores anteriores de cliente y forma de pago
    ClienteAnt = Text1(3).Text
    FormaPagoAnt = Text1(5).Tag
    
    'Quan modifique pose en la F.Modificación la data actual
    PonerFoco Text1(3)
End Sub

Private Sub BotonRectificar()
    
    Set frmList = New frmListado
    'Añadiremos el boton de aceptar y demas objetos para insertar
    frmList.CadTag = Text1(0).Text & "|" & Text1(1).Text & "|" & Text1(2).Text & "|" & Text1(3).Text & "|" & text2(3).Text & "|" & Format(Check1(1).Value, "0") & "|"
    frmList.OpcionListado = 12
    frmList.Show vbModal

End Sub

Private Sub BotonEliminar()
Dim Cad As String

    On Error GoTo EEliminar

    'Ciertas comprobaciones
    If Data1.Recordset.EOF Then Exit Sub
    
    'El registre de codi 0 no es pot Modificar ni Eliminar
    If EsCodigoCero(CStr(Data1.Recordset.Fields(1).Value), FormatoCampo(Text1(1))) Then Exit Sub

    Cad = "¿Seguro que desea eliminar la factura?"
    Cad = Cad & vbCrLf & "Nº: " & Format(Data1.Recordset!numfactu, FormatoCampo(Text1(1)))
    Cad = Cad & vbCrLf & "Fecha: " & Data1.Recordset.Fields("fecfactu")
    
    'Borramos
    If MsgBox(Cad, vbQuestion + vbYesNo) = vbYes Then
        'Hay que eliminar
        On Error GoTo EEliminar
        Screen.MousePointer = vbHourglass
        NumRegElim = Data1.Recordset.AbsolutePosition
        If Not Eliminar Then
            Screen.MousePointer = vbDefault
            Exit Sub
        ElseIf SituarDataTrasEliminar(Data1, NumRegElim) Then
            PonerCampos
        Else
            LimpiarCampos
            'Poner los grid sin apuntar a nada
            'LimpiarDataGrids
            PonerModo 0
        End If
    End If
    Screen.MousePointer = vbDefault
    
EEliminar:
    Screen.MousePointer = vbDefault
    If Err.Number > 0 Then MuestraError Err.Number, "Eliminar Factura", Err.Description
End Sub

Private Sub PonerCampos()
Dim I As Integer
Dim ImporteVale As Currency




    If Data1.Recordset.EOF Then Exit Sub
    PonerCamposForma2 Me, Data1, 1 'opcio=1: pone el formato o los campos de la cabecera
    
    For I = 0 To DataGridAux.Count - 1
        CargaGrid I, True
    Next I
    
    'Recuperar Descripciones de los campos de Codigo
    '--------------------------------------------------
    text2(3).Text = PonerNombreDeCod(Text1(3), "ssocio", "nomsocio")
    text2(4).Text = PonerNombreDeCod(Text1(4), "scoope", "nomcoope")
    text2(5).Text = PonerNombreDeCod(Text1(5), "sforpa", "nomforpa")
    text2(26).Text = PonerNombreDeCod(Text1(26), "departamento", "nomdepar")
    
    '[Monica]28/12/2015: ponemos el importe del vale
    ImporteVale = DevuelveValor("select sum(coalesce(importevale,0)) from slhfac where " & ObtenerWhereCab(False))
    Text1(25).Text = ""
    If ImporteVale <> 0 Then
        Text1(25).Text = Format(ImporteVale, "###,###,##0.00")
    End If
    
    

    '-- Esto permanece para saber donde estamos
    lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
    PonerModoOpcionesMenu (Modo)
End Sub

Private Sub cmdCancelar_Click()
Dim I As Integer
Dim V

    Select Case Modo
        Case 1, 3 'Busqueda, Insertar
                LimpiarCampos
                If Data1.Recordset.EOF Then
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
        
        Case 5 'LINEAS
            Select Case ModoLineas
'                Case 1 'afegir llinia
'                    ModoLineas = 0
'                    DataGridAux(NumTabMto).AllowAddNew = False
''                    SituarTab (NumTabMto)
'                    'If BLOQUEADesdeFormulario2(Me, Data1, 1) Then BotonModificar  'Modificar
'                    LLamaLineas NumTabMto, ModoLineas 'ocultar txtAux
'                    'If DataGridAux(NumTabMto).Enabled Then DataGridAux(NumTabMto).SetFocus
'                    DataGridAux(NumTabMto).Enabled = True
'                    DataGridAux(NumTabMto).SetFocus
'
'                    If Not AdoAux(NumTabMto).Recordset.EOF Then
'                        AdoAux(NumTabMto).Recordset.MoveFirst
'                    End If

                Case 2 'modificar llinies
                    ModoLineas = 0
                    PonerModo 4
                    If Not AdoAux(NumTabMto).Recordset.EOF Then
                        V = AdoAux(NumTabMto).Recordset.Fields(3) 'el 1 es el nº de llinia
                        AdoAux(NumTabMto).Recordset.Find (AdoAux(NumTabMto).Recordset.Fields(1).Name & " =" & V)
                    End If
                    LLamaLineas NumTabMto, ModoLineas 'ocultar txtAux
            End Select
            
            PosicionarData "letraser = '" & Text1(0).Text & "' and numfactu = " & Text1(1).Text & " and fecfactu = " & DBSet(Text1(2).Text, "F")
            
'            If Not AdoAux(NumTabMto).Recordset.EOF Then
'                DataGridAux_RowColChange NumTabMto, 1, 1
'            Else
'                LimpiarCamposFrame NumTabMto
'            End If
    End Select
End Sub

Private Function DatosOk() As Boolean
Dim b As Boolean
Dim Datos As String
Dim SQL As String

    On Error GoTo EDatosOK

    DatosOk = False
    b = CompForm2(Me, 1) 'Comprobar formato datos ok
    If Not b Then Exit Function
    
    ' en caso de que haya contabilidad
    ' comprobamos si el cliente tiene cuenta contable existente en la contabilidad
    If vParamAplic.NumeroConta <> 0 And Check1(1).Value <> 0 Then
        SQL = ""
        SQL = DevuelveDesdeBD("codmacta", "ssocio", "codsocio", Text1(3).Text, "N")
        If SQL = "" Then
            MsgBox "El cliente no tiene cuenta contable asociada.", vbExclamation
            Exit Function
        Else
            Datos = ""
            Datos = DevuelveDesdeBDNew(cConta, "cuentas", "codmacta", "codmacta", SQL, "T")
            If Datos = "" Then
                MsgBox "La cuenta contable asociada al cliente no está dada de alta en contabilidad. Revise.", vbExclamation
                Exit Function
            End If
        End If
    End If
    
    'Comprobamos que la suma de importes de las lineas es igual al total de la factura
    Datos = SumaLineas("")
    
    If CCur(Datos) > (CCur(Text1(18).Text) + CCur(Text1(19).Text)) Then
        MsgBox "La suma de los importes de lineas es mayor que el total de la factura!!!", vbExclamation
    ElseIf CCur(Datos) < CCur(Text1(18).Text) Then
        MsgBox "La suma de los importes de lineas es menor que el total de la factura!!!", vbExclamation
    End If
         
    DatosOk = b
EDatosOK:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Function

Private Sub PosicionarData(Cad As String)
'Dim cad As String
Dim Indicador As String
    
  '  cad = ""
    If SituarDataMULTI(Data1, Cad, Indicador) Then
        If ModoLineas <> 1 Then
            PonerModo 2
        End If
       
       lblIndicador.Caption = Indicador
    Else
       LimpiarCampos
       'Poner los grid sin apuntar a nada
       'LimpiarDataGrids
       PonerModo 0
    End If
End Sub

Private Function Eliminar() As Boolean
Dim vWhere As String

    On Error GoTo FinEliminar
        
    Conn.BeginTrans
    vWhere = ObtenerWhereCab(True)

    'Eliminar las Lineas de facturas de proveedor
    Conn.Execute "DELETE FROM slhfac " & vWhere
    
    'Eliminar la CABECERA
    Conn.Execute "Delete from " & NomTabla & vWhere
               
FinEliminar:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Eliminar"
        Conn.RollbackTrans
        Eliminar = False
    Else
        Conn.CommitTrans
        Eliminar = True
    End If
End Function

Private Sub Text1_GotFocus(Index As Integer)
    kCampo = Index
    ConseguirFoco Text1(Index), Modo
End Sub

Private Sub Text1_LostFocus(Index As Integer)
Dim Cad As String, Datos As String
Dim Suma As Currency

    If Not PerderFocoGnral(Text1(Index), Modo) Then Exit Sub
    
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub
    
    Select Case Index
        Case 1, 21 'Nº factura
            If Text1(Index).Text <> "" Then FormateaCampo Text1(Index)
            
        Case 2, 23 'Fecha
            If Text1(Index).Text <> "" Then PonerFormatoFecha Text1(Index)
            
        Case 3 'Cliente
            If Text1(Index).Text <> "" Then
                If PonerFormatoEntero(Text1(3)) Then
                    text2(Index).Text = PonerNombreDeCod(Text1(Index), "ssocio", "nomsocio", "codsocio", "N")
                    If text2(Index).Text = "" Then
                        Cad = "No existe el Cliente: " & Text1(Index).Text & vbCrLf
                        Cad = Cad & "¿Desea crearlo?" & vbCrLf
                        If MsgBox(Cad, vbQuestion + vbYesNo) = vbYes Then
                            Set frmcli = New frmManClien
                            frmcli.DatosADevolverBusqueda = "0|1|"
                            Text1(Index).Text = ""
                            TerminaBloquear
                            frmcli.Show vbModal
                            Set frmcli = Nothing
                            If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Data1, 1
                        Else
                            Text1(Index).Text = ""
                        End If
                        PonerFoco Text1(Index)
                    End If
                Else
                    text2(Index).Text = ""
                End If
                'recuperar el Colectivo
                If Modo = 1 Then Exit Sub
                 Datos = DevuelveDesdeBDNew(cPTours, "ssocio", "codcoope", "codsocio", Text1(3).Text, "N")
                 If Datos <> "" Then
                     Text1(4).Text = Datos
                     FormateaCampo Text1(4)
                     Text1_LostFocus (4)
                 Else
                     Text1(4).Text = ""
                     text2(4).Text = ""
                 End If
            End If
            
        Case 4 'Colectivo
            If PonerFormatoEntero(Text1(4)) Then
                text2(Index).Text = PonerNombreDeCod(Text1(Index), "scoope", "nomcoope", "codcoope", "N")
                If text2(Index).Text = "" Then
                    Cad = "No existe el Colectivo: " & Text1(Index).Text & vbCrLf
                    Cad = Cad & "¿Desea crearlo?" & vbCrLf
                    If MsgBox(Cad, vbQuestion + vbYesNo) = vbYes Then
                        Set frmCol = New frmManCoope
                        frmCol.DatosADevolverBusqueda = "0|1|"
                        Text1(Index).Text = ""
                        TerminaBloquear
                        frmCol.Show vbModal
                        Set frmCol = Nothing
                        If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Data1, 1
                    Else
                        Text1(Index).Text = ""
                    End If
                    PonerFoco Text1(Index)
                End If
            Else
                text2(Index).Text = ""
            End If
        
        Case 5 'Forma pago
            If PonerFormatoEntero(Text1(5)) Then
                text2(Index).Text = PonerNombreDeCod(Text1(Index), "sforpa", "nomforpa", "codforpa", "N")
                If text2(Index).Text = "" Then
                    Cad = "No existe la Forma de Pago: " & Text1(Index).Text & vbCrLf
                    Cad = Cad & "¿Desea crearla?" & vbCrLf
                    If MsgBox(Cad, vbQuestion + vbYesNo) = vbYes Then
                        Set frmFPa = New frmManFpago
                        frmFPa.DatosADevolverBusqueda = "0|1|"
                        Text1(Index).Text = ""
                        TerminaBloquear
                        frmFPa.Show vbModal
                        Set frmFPa = Nothing
                        If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Data1, 1
                    Else
                        Text1(Index).Text = ""
                    End If
                    PonerFoco Text1(Index)
                End If
            Else
                text2(Index).Text = ""
            End If
            
        Case 6, 9, 10, 13, 14, 17 'IMPORTES Base, IVA
            
'            If PonerFormatoDecimal(Text1(Index), 1) Then
'                CalcularTotalFactura Index
'            End If
            
       Case 7, 11, 15 'cod. IVA
           If Text1(Index).Text = "" Then
              Text1(Index + 1).Text = ""
           Else
                Datos = DevuelveDesdeBDNew(cConta, "tiposiva", "porceiva", "codigiva", Text1(Index).Text, "N")
                If Datos = "Error" Then
                    PonerFoco Text1(Index)
                    Exit Sub
                End If
                If Datos = "" Then
                    MsgBox "No existe el código de IVA: " & Text1(Index).Text, vbExclamation
                    Text1(Index).Text = ""
                    PonerFoco Text1(Index)
                Else
                    FormateaCampo Text1(Index)
                    Text1(Index + 1).Text = Datos
                End If
            End If
'            CalcularTotalFactura
              
    End Select
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = teclaBuscar Then
        If Modo = 1 Or Modo = 3 Or Modo = 4 Then
            Select Case Index
                Case 2: KEYFecha KeyAscii, 2
                Case 3: KEYBusqueda KeyAscii, 0
                Case 4: KEYBusqueda KeyAscii, 1
                Case 5: KEYBusqueda KeyAscii, 2
                Case 7: KEYBusqueda KeyAscii, 3
                Case 11: KEYBusqueda KeyAscii, 4
                Case 15: KEYBusqueda KeyAscii, 5
               ' Case 1: KEYFecha KeyAscii, 1
            End Select
        End If
    Else
        KEYpress KeyAscii
    End If
End Sub

Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'Avanzar/Retroceder los campos con las flechas de desplazamiento del teclado.
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
    btnBuscar_Click (indice)
End Sub

'************* LLINIES: ****************************
Private Sub ToolAux_ButtonClick(Index As Integer, ByVal Button As MSComctlLib.Button)

    If vParamAplic.NumeroConta <> 0 And _
       Not FacturaModificable(Text1(0).Text, Text1(1).Text, Text1(2).Text, Check1(1).Value) Then Exit Sub
    
    Select Case Button.Index
'        Case 1
''            TerminaBloquear
'            BotonAnyadirLinea Index
        Case 2
'            TerminaBloquear
            BotonModificarLinea Index
        Case 3
'            TerminaBloquear
            BotonEliminarLinea Index
            If Modo = 4 Then
                If BLOQUEADesdeFormulario2(Me, Data1, 1) Then BotonModificar
            End If
        Case 6 'Imprimir
'            BotonImprimirLinea Index
    End Select
End Sub

Private Sub BotonEliminarLinea(Index As Integer)
Dim SQL As String
Dim Eliminar As Boolean

    On Error GoTo Error2

    ModoLineas = 3 'Posem Modo Eliminar Llínia

    If Modo = 4 Then 'Modificar Capçalera
        cmdAceptar_Click
        If ModoLineas = 0 Then Exit Sub
    End If

    NumTabMto = Index
    PonerModo 5

    If AdoAux(Index).Recordset.EOF Then Exit Sub
    If Not SepuedeBorrar(Index) Then Exit Sub
    
    If AdoAux(Index).Recordset.RecordCount = 1 Then
        MsgBox "No se puede borrar un única línea de factura, elimine la factura completa", vbExclamation
        PonerModo 2
        Exit Sub
    End If
    
    
    Eliminar = False

    Select Case Index
        Case 0 'lineas de factura
            SQL = "¿Seguro que desea eliminar la línea?"
            SQL = SQL & vbCrLf & "Nº línea: " & Format(DBLet(AdoAux(Index).Recordset!NumLinea), FormatoCampo(txtAux(3)))
            SQL = SQL & vbCrLf & "Albaran: " & DBLet(AdoAux(Index).Recordset!numalbar) '& "  " & txtAux(4).Text
            If MsgBox(SQL, vbQuestion + vbYesNo) = vbYes Then
                NumRegElim = AdoAux(Index).Recordset.AbsolutePosition
                Eliminar = True
                SQL = "DELETE FROM slhfac"
                SQL = SQL & ObtenerWhereCab(True) & " AND numlinea= " & AdoAux(Index).Recordset!NumLinea
            End If
    End Select

    If Eliminar Then
        TerminaBloquear
'        conn.Execute sql
        CadenaBorrado = SQL
        '16022007
        If BLOQUEADesdeFormulario2(Me, Data1, 1) Then
                ModificaImportes = True
                BotonModificar
                cmdAceptar_Click
        End If
        'EliminarLinea
        
        
        'antes estaba debajo de situardata
        CargaGrid Index, True
        SituarDataTrasEliminar AdoAux(Index), NumRegElim, True
    End If

    ModoLineas = 0
    PosicionarData "letraser = '" & Text1(0).Text & "' and numfactu = " & Text1(1).Text & " and fecfactu = " & DBSet(Text1(2).Text, "F")

    Exit Sub
Error2:
    Screen.MousePointer = vbDefault
    MuestraError Err.Number, "Eliminando linea", Err.Description
End Sub

'Private Sub BotonAnyadirLinea(Index As Integer)
'Dim NumF As String
'Dim vWhere As String, vTabla As String
'Dim anc As Single
'Dim i As Integer
'Dim SumLin As Currency
'
'    'Si no estaba modificando lineas salimos
'    ' Es decir, si estaba insertando linea no podemos hacer otra cosa
'    'If ModificaLineas = 2 Then Exit Sub
'    ModoLineas = 1 'Ponemos Modo Añadir Linea
'
'    If (Modo = 3) Or (Modo = 4) Then 'Insertar o Modifcar Cabecera
'        cmdAceptar_Click
'        'No se ha insertado la cabecera
'        If ModoLineas = 0 Then Exit Sub
'    End If
'
'    NumTabMto = Index
'    PonerModo 5
''    If b Then BloquearText1 Me, 4 'Si viene de Insertar Cabecera no bloquear los Text1
'
'
'    'Obtener el numero de linea ha insertar
'    Select Case Index
'        Case 0: vTabla = "slhfac"
'    End Select
'    'Obtener el sig. nº de linea a insertar
'    vWhere = ObtenerWhereCab(False)
'    NumF = SugerirCodigoSiguienteStr(vTabla, "numlinea", vWhere)
'
'    'Situamos el grid al final
'    AnyadirLinea DataGridAux(Index), AdoAux(Index)
'
'    anc = DataGridAux(Index).Top
'    If DataGridAux(Index).Row < 0 Then
'        anc = anc + 210
'    Else
'        anc = anc + DataGridAux(Index).RowTop(DataGridAux(Index).Row) + 5
'    End If
'
'    LLamaLineas Index, ModoLineas, anc
'
'    Select Case Index
'        Case 0 'lineas factura
'            txtAux(0).Text = Text1(0).Text 'serie
'            txtAux(1).Text = Text1(1).Text 'factura
'            txtAux(2).Text = Text1(2).Text 'fecha
'            txtAux(3).Text = NumF 'numlinea
'            FormateaCampo txtAux(3)
'            For i = 4 To 12
'                txtAux(i).Text = ""
'            Next i
'
'            'desbloquear la linea (se bloquea al añadir)
'            BloquearTxt txtAux(3), False
'            PonerFoco txtAux(4)
'    End Select
'End Sub

Private Sub BotonModificarLinea(Index As Integer)
    Dim anc As Single
    Dim I As Integer
    Dim J As Integer
    
    If AdoAux(Index).Recordset.EOF Then Exit Sub
    If AdoAux(Index).Recordset.RecordCount < 1 Then Exit Sub
    
    ModoLineas = 2 'Modificar llínia
    
    If Modo = 4 Then 'Modificar Cabecera
        cmdAceptar_Click
        If ModoLineas = 0 Then Exit Sub
    End If
    
    
    NumTabMto = Index
    PonerModo 5
    
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

    Select Case Index
        Case 0 'lineas de factura
            For J = 0 To 9
                txtAux(J).Text = DataGridAux(Index).Columns(J).Text
            Next J
'            txtAux(6).Text = DataGridAux(Index).Columns(j).Text
'            txtAux(7).Text = DataGridAux(Index).Columns(j + 1).Text
            
            txtAux2(0).Text = DataGridAux(Index).Columns(10).Text
            For J = 10 To 12
                txtAux(J) = DataGridAux(Index).Columns(J + 1).Text
            Next J
            
'            txtAux2(6).Text = DataGridAux(Index).Columns(j + 1).Text
    
    End Select
    
    LLamaLineas Index, ModoLineas, anc
   
    Select Case Index
        Case 0 'lineas de factura
            PonerFoco txtAux(4)
    End Select
End Sub

Private Sub LLamaLineas(Index As Integer, xModo As Byte, Optional alto As Single)
Dim jj As Integer
Dim b As Boolean

    On Error GoTo ELLamaLin

    DeseleccionaGrid DataGridAux(Index)
    
    b = (xModo = 1 Or xModo = 2) 'Insertar o Modificar Lineas
    Select Case Index
        Case 0 'lineas de factura
            For jj = 4 To 12
                txtAux(jj).Top = alto
                txtAux(jj).visible = b
            Next jj
            txtAux2(0).Top = alto
            txtAux2(0).visible = b
            Me.btnBuscar(0).Top = alto
            Me.btnBuscar(0).visible = b
'            Me.btnBuscar(1).Top = alto
'            Me.btnBuscar(1).visible = b
    End Select
    
ELLamaLin:
    Err.Clear
End Sub

Private Sub txtAux_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = teclaBuscar Then
        If Modo = 5 And (ModoLineas = 1 Or ModoLineas = 2) Then
            Select Case Index
                Case 5: KEYBusquedaLin KeyAscii, 0
                Case 6: KEYBusquedaLin KeyAscii, 1
            End Select
        End If
    Else
        KEYpress KeyAscii
    End If
End Sub

Private Sub txtAux_LostFocus(Index As Integer)
Dim cadMen As String

    txtAux(Index).Text = Trim(txtAux(Index).Text)

    Select Case Index
        Case 4 ' albaran
            txtAux(Index).Text = UCase(txtAux(Index).Text)
            
        Case 5 ' fecha de albaran
            PonerFormatoFecha txtAux(Index)

        Case 6 ' hora
            PonerFormatoHora txtAux(Index)

        Case 7 ' turno
            If Not EsNumerico(txtAux(Index).Text) Then
                MsgBox "El turno debe ser numérico.", vbExclamation
                On Error Resume Next
                txtAux(Index).Text = ""
                PonerFoco txtAux(Index)
                Exit Sub
            End If
            FormateaCampo txtAux(Index)
            
        Case 8 ' tarjeta
            If Not EsNumerico(txtAux(Index).Text) Then
                MsgBox "El número de tarjeta debe ser numérico.", vbExclamation
                On Error Resume Next
                txtAux(Index).Text = ""
                PonerFoco txtAux(Index)
                Exit Sub
            End If
            FormateaCampo txtAux(Index)

        Case 9 ' articulo
            If PonerFormatoEntero(txtAux(Index)) Then
                txtAux2(0).Text = PonerNombreDeCod(txtAux(Index), "sartic", "nomartic", "codartic", "N")
                If txtAux2(0).Text = "" Then
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
                    Else
                        txtAux(Index).Text = ""
                    End If
                    PonerFoco txtAux(Index)
                End If
            Else
                txtAux2(0).Text = ""
            End If

        Case 10 ' cantidad
           If Not EsNumerico(txtAux(Index).Text) Then
                MsgBox "La cantidad debe ser numérica.", vbExclamation
                On Error Resume Next
                txtAux(Index).Text = ""
                PonerFoco txtAux(Index)
                Exit Sub
            End If
            'Es numerico
            PonerFormatoDecimal txtAux(Index), 2
        Case 11 ' precio
           If Not EsNumerico(txtAux(Index).Text) Then
                MsgBox "El Precio debe ser numérico.", vbExclamation
                On Error Resume Next
                txtAux(Index).Text = ""
                PonerFoco txtAux(Index)
                Exit Sub
            End If
            'Es numerico
            PonerFormatoDecimal txtAux(Index), 2

        Case 12 'Importe
           If Trim(txtAux(Index).Text) = "" Then
                PonerFocoBtn Me.cmdAceptar
                Exit Sub
           End If
           If Not EsNumerico(txtAux(Index).Text) Then
                MsgBox "El Importe debe ser numérico.", vbExclamation
                On Error Resume Next
                txtAux(Index).Text = ""
                PonerFoco txtAux(Index)
                Exit Sub
            End If
            'Es numerico
            PonerFormatoDecimal txtAux(Index), 3
            PonerFocoBtn Me.cmdAceptar
    End Select
    
    CalcularImporteNue txtAux(10), txtAux(11), txtAux(12), Index - 10
    
End Sub

Private Sub txtAux_GotFocus(Index As Integer)
    ConseguirFocoLin txtAux(Index)
End Sub

Private Function DatosOkLlin(nomframe As String) As Boolean
Dim b As Boolean
Dim SumLin As Currency
    
    On Error GoTo EDatosOKLlin

    DatosOkLlin = False
        
    b = CompForm2(Me, 2, nomframe) 'Comprobar formato datos ok
    If Not b Then Exit Function
    
' ### [Monica] 29/09/2006
' he quitado la parte de comprobar la suma de lineas
'    'Comprobar que el Importe del total de las lineas suma el total o menos de la factura
'    SumLin = CCur(SumaLineas(txtAux(4).Text))
'
'    'Le añadimos el importe de linea que vamos a insertar
'    SumLin = SumLin + CCur(txtAux(7).Text)
'
'    'comprobamos que no sobrepase el total de la factura
'    If SumLin > CCur(Text1(18).Text) Then
'        MsgBox "La suma del importe de las lineas no puede ser superior al total de la factura.", vbExclamation
'        b = False
'    End If
    
    DatosOkLlin = b
    
EDatosOKLlin:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Function

Private Function SepuedeBorrar(ByRef Index As Integer) As Boolean

    SepuedeBorrar = False
    If AdoAux(Index).Recordset.EOF Then Exit Function

    SepuedeBorrar = True
End Function

Private Sub CargaGrid(Index As Integer, enlaza As Boolean)
Dim b As Boolean
Dim tots As String

    On Error GoTo ECarga

    'b = DataGridAux(Index).Enabled
    'DataGridAux(Index).Enabled = False
    
    tots = MontaSQLCarga(Index, enlaza)
    
    CargaGridGnral Me.DataGridAux(Index), Me.AdoAux(Index), tots, PrimeraVez
    
    Select Case Index
        Case 0 'lineas de factura
            'si es visible|control|tipo campo|nombre campo|ancho control|formato campo|
            tots = "N||||0|;N||||0|;N||||0|;N||||0|;"
            tots = tots & "S|txtAux(4)|T|Albaran|900|;S|txtAux(5)|T|Fecha|1000|;S|txtAux(6)|T|Hora|650|;"
            tots = tots & "S|txtAux(7)|T|Tur.|400|;S|txtAux(8)|T|Tarjeta|1400|;S|txtAux(9)|T|Articulo|800|;"
            tots = tots & "S|btnBuscar(0)|B|||;S|txtAux2(0)|T|Denominación|2030|;S|txtAux(10)|T|Cantidad|1000|;"
            tots = tots & "S|txtAux(11)|T|Precio|900|;S|txtAux(12)|T|Importe|1100|;"
            arregla tots, DataGridAux(Index), Me
'           DataGridAux(Index).Columns(6).Alignment = dbgCenter
'           DataGridAux(Index).Columns(9).Alignment = dbgRight
    End Select
    
    DataGridAux(Index).ScrollBars = dbgAutomatic
    
ECarga:
    If Err.Number <> 0 Then MuestraError Err.Number, "Cargando datos grid: " & DataGridAux(Index).Tag, Err.Description
End Sub

'Private Sub InsertarLinea()
''Inserta registro en las tablas de Lineas: provbanc, provdpto
'Dim nomFrame As String
'Dim b As Boolean
'
'    On Error Resume Next
'
'    Select Case NumTabMto
'        Case 0: nomFrame = "FrameAux0" 'lineas de factura
'    End Select
'
'    If DatosOkLlin(nomFrame) Then
'        TerminaBloquear
'        If InsertarDesdeForm2(Me, 2, nomFrame) Then
'            b = BLOQUEADesdeFormulario2(Me, Data1, 1)
'            CargaGrid NumTabMto, True
'            If b Then BotonAnyadirLinea NumTabMto
'        End If
'    End If
'End Sub

Private Sub ModificarLinea()
'Modifica registro en las tablas de Lineas: provbanc, provdpto
Dim nomframe As String
Dim V As Currency

' variables para el recalculo de iva y totales
    Dim I As Integer
    Dim Imptot(2)
    Dim Tipiva(2)
    Dim Impbas(2) As Currency
    Dim ImpIva(2) As Currency
    Dim PorIva(2) As Currency
    Dim TotFac As Currency
    Dim totimp As Currency
    Dim totimpSigaus As Currency



    On Error GoTo EModificarLin

    Select Case NumTabMto
        Case 0: nomframe = "FrameAux0" 'lineas de factura
    End Select
    
    If DatosOkLlin(nomframe) Then
        TerminaBloquear
'        conn.BeginTrans
        If ModificaDesdeFormulario2(Me, 2, nomframe) Then
            If BLOQUEADesdeFormulario2(Me, Data1, 1) Then
            
            ' ### [Monica] 29/09/2006
            ' he quitado el boton modificar para recalcular bases e iva
            
            'BotonModificar
                

                
            End If
            V = AdoAux(NumTabMto).Recordset.Fields(1) 'el 2 es el nº de llinia
            ModoLineas = 0
            CargaGrid NumTabMto, True
'            SituarTab (NumTabMto)
            DataGridAux(NumTabMto).SetFocus
            AdoAux(NumTabMto).Recordset.Find (AdoAux(NumTabMto).Recordset.Fields(1).Name & " =" & V)
            
'            ' ### [Monica] 29/09/2006
'            ' añadido el tema de de recalculo de bases
            RecalculoBasesIvaFactura AdoAux(0).Recordset, Imptot, Tipiva, Impbas, ImpIva, PorIva, TotFac, totimp, totimpSigaus

            
            '13/02/2007 iniacializo los txt
            For I = 0 To 2
                Text1(7 + (4 * I)).Text = ""
                Text1(6 + (4 * I)).Text = ""
                Text1(8 + (4 * I)).Text = ""
                Text1(9 + (4 * I)).Text = ""
            Next I
            
            '13/02/2007 he añadido las condiciones del for antes solo estaban las sentencias
            For I = 0 To 2
                 If Tipiva(I) <> 0 Then Text1(7 + (4 * I)).Text = Tipiva(I)
                 If Impbas(I) <> 0 Then Text1(6 + (4 * I)).Text = Impbas(I)
                 If PorIva(I) <> 0 Then Text1(8 + (4 * I)).Text = PorIva(I)
                 If ImpIva(I) <> 0 Then Text1(9 + (4 * I)).Text = ImpIva(I)
                 'TotFac = Impbas(i) + impiva(i)
            Next I
        
            Text1(19).Text = totimp
            Text1(20).Text = totimpSigaus
            Text1(18).Text = TotFac
            
            PonerFormatoDecimal Text1(18), 1
            PonerFormatoDecimal Text1(19), 1
            PonerFormatoDecimal Text1(20), 1
            
            If BLOQUEADesdeFormulario2(Me, Data1, 1) Then
'                Modo = 4
'                PonerModo Modo
'                ClienteAnt = Text1(3).Text
'                FormaPagoAnt = Text1(5).Text
                ModificaImportes = True
                BotonModificar
                cmdAceptar_Click

            End If

            LLamaLineas NumTabMto, 0
        End If
    End If
    Exit Sub
    
EModificarLin:
    If Err.Number <> 0 Then MuestraError Err.Number, "Modificar Linea", Err.Description
End Sub

Private Function ObtenerWhereCab(conW As Boolean) As String
Dim vWhere As String
    vWhere = ""
    If conW Then vWhere = " WHERE "
    vWhere = vWhere & " letraser='" & Text1(0).Text & "'"
    vWhere = vWhere & " AND numfactu= " & Text1(1).Text & " AND fecfactu= '" & Format(Text1(2).Text, FormatoFecha) & "'"
    ObtenerWhereCab = vWhere
End Function



Private Function SumaLineas(NumLin As String) As String
'Al Insertar o Modificar linea sumamos todas las lineas excepto la que estamos
'Insertando o modificando que su valor sera el del txtaux(4).text
'En el DatosOK de la factura sumamos todas las lineas
Dim SQL As String
Dim Rs As ADODB.Recordset
Dim SumLin As Currency

    SumLin = 0
    Select Case Tipo
        Case 0:
            SQL = "SELECT SUM(implinea) FROM slhfac "
        Case 1:
            SQL = "SELECT SUM(implinea) FROM slhfacr "
        Case 2:
            SQL = "SELECT SUM(implinea) FROM slhfac1 "
    End Select
    SQL = SQL & ObtenerWhereCab(True)
    If NumLin <> "" Then SQL = SQL & " AND numlinea<>" & DBSet(txtAux(4).Text, "N") 'numlinea
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not Rs.EOF Then
        'En SumLin tenemos la suma de las lineas ya insertadas
        SumLin = CCur(DBLet(Rs.Fields(0), "N"))
    End If
    Rs.Close
    Set Rs = Nothing
    SumaLineas = CStr(SumLin)
End Function

'' ### [DavidV] 26/04/2006: Activar/desactivar la rueda del ratón.
'Private Sub DataGridAux_GotFocus(Index As Integer)
'  WheelHook DataGridAux(Index)
'End Sub
'Private Sub DataGridAux_LostFocus(Index As Integer)
'  WheelUnHook
'End Sub


Private Function FacturaModificable(letraser As String, numfactu As String, fecfactu As String, Contabil As String) As Boolean

    FacturaModificable = False
    
    If Contabil = 0 Then
        FacturaModificable = True
    Else
        ' si la factura esta contabilizada tenemos que ver si en la contabilidad esta contabilizada y
        ' si en la tesoreria esta remesada o cobrada en estos casos la factura no puede ser modificada
        If FacturaContabilizada(letraser, numfactu, Year(CDate(fecfactu))) Then
            MsgBox "Factura contabilizada en la Contabilidad, no puede modificarse ni eliminarse."
            Exit Function
        End If
        
        If FacturaRemesada(letraser, numfactu, fecfactu) Then
            MsgBox "Factura Remesada, no puede modificarse ni eliminarse."
            Exit Function
        End If
        
        If FacturaCobrada(letraser, numfactu, fecfactu) Then
            MsgBox "Factura Cobrada, no puede modificarse ni eliminarse."
            Exit Function
        End If
           
        'solo se puede modificar la factura si no esta contabilizada
        If FactContabilizada2(letraser, numfactu, fecfactu) Then
            TerminaBloquear
            Exit Function
        End If
           
           
           
           
           
        FacturaModificable = True
    End If

End Function

'VRS:2.0.1(3)
Private Sub LlamarImprimir()

    With frmImprimir
        'Nuevo. Febrero 2010
        .outClaveNombreArchiv = Text1(0).Text & Format(Text1(1).Text, "000")
        .outCodigoCliProv = Text1(3).Text
        .outTipoDocumento = 100
        
        
        
        .FormulaSeleccion = cadFormula
        .OtrosParametros = cadParam
        .NumeroParametros = 2
        .SoloImprimir = False
        .Titulo = cadTitulo
        .NombreRPT = cadNombreRPT
        .EnvioEMail = False
        .Opcion = OpcionListado '0
        .Show vbModal
    End With
End Sub

Private Sub CambiarTags(Tipo As Byte)
Dim Txt As TextBox

    For Each Txt In Me.Text1
        Select Case Tipo
            Case 0:
                Txt.Tag = Replace(Txt.Tag, "schfacr", "schfac")
            Case 1:
                Txt.Tag = Replace(Txt.Tag, "schfac", "schfacr")
            Case 2:
                Txt.Tag = Replace(Txt.Tag, "schfac", "schfac1")
        End Select
    Next Txt
    For Each Txt In Me.txtAux
        Select Case Tipo
            Case 0:
                Txt.Tag = Replace(Txt.Tag, "slhfacr", "slhfac")
            Case 1:
                Txt.Tag = Replace(Txt.Tag, "slhfac", "slhfacr")
            Case 2:
                Txt.Tag = Replace(Txt.Tag, "slhfac", "slhfac1")
        End Select
    Next Txt


End Sub


Private Sub ActivarFrameCobros()
Dim obj As Object

For Each obj In Me
    If TypeOf obj Is Frame Then
        If obj.Name = "FrameCobros" Then
            
            
        End If
        
    End If
Next obj

End Sub


Private Sub EliminarLinea()
Dim nomframe As String
Dim V As Currency

' variables para el recalculo de iva y totales
    Dim I As Integer
    Dim Imptot(2)
    Dim Tipiva(2)
    Dim Impbas(2) As Currency
    Dim ImpIva(2) As Currency
    Dim PorIva(2) As Currency
    Dim TotFac As Currency
    Dim totimp As Currency
    Dim totimpSigaus As Currency



    On Error GoTo EEliminarLin

    Select Case NumTabMto
        Case 0: nomframe = "FrameAux0" 'lineas de factura
    End Select
    
    TerminaBloquear
'        conn.BeginTrans
    If BLOQUEADesdeFormulario2(Me, Data1, 1) Then
            
            ' ### [Monica] 29/09/2006
            ' he quitado el boton modificar para recalcular bases e iva
            
            'BotonModificar
                

                
            End If
            V = AdoAux(NumTabMto).Recordset.Fields(1) 'el 2 es el nº de llinia
            ModoLineas = 0
            CargaGrid NumTabMto, True
'            SituarTab (NumTabMto)
            DataGridAux(NumTabMto).SetFocus
            AdoAux(NumTabMto).Recordset.Find (AdoAux(NumTabMto).Recordset.Fields(1).Name & " =" & V)
            
'            ' ### [Monica] 29/09/2006
'            ' añadido el tema de de recalculo de bases
            RecalculoBasesIvaFactura AdoAux(0).Recordset, Imptot, Tipiva, Impbas, ImpIva, PorIva, TotFac, totimp, totimpSigaus

            
            '13/02/2007 iniacializo los txt
            For I = 0 To 2
                Text1(7 + (4 * I)).Text = ""
                Text1(6 + (4 * I)).Text = ""
                Text1(8 + (4 * I)).Text = ""
                Text1(9 + (4 * I)).Text = ""
            Next I
            
            '13/02/2007 he añadido las condiciones del for antes solo estaban las sentencias
            For I = 0 To 2
                 If Tipiva(I) <> 0 Then Text1(7 + (4 * I)).Text = Tipiva(I)
                 If Impbas(I) <> 0 Then Text1(6 + (4 * I)).Text = Impbas(I)
                 If PorIva(I) <> 0 Then Text1(8 + (4 * I)).Text = PorIva(I)
                 If ImpIva(I) <> 0 Then Text1(9 + (4 * I)).Text = ImpIva(I)
                 'TotFac = Impbas(i) + impiva(i)
            Next I
            Text1(19).Text = totimp
            Text1(20).Text = totimpSigaus
            Text1(18).Text = TotFac
            
            PonerFormatoDecimal Text1(18), 1
            PonerFormatoDecimal Text1(19), 1
            PonerFormatoDecimal Text1(20), 1
            
'            If BLOQUEADesdeFormulario2(Me, Data1, 1) Then
'                ModificaImportes = True
'                BotonModificar
'                cmdAceptar_Click
'            End If

            LLamaLineas NumTabMto, 0
    Exit Sub
    
EEliminarLin:
    If Err.Number <> 0 Then MuestraError Err.Number, "Eliminar Linea", Err.Description
End Sub

Private Function FactContabilizada2(letraser As String, numfactu As String, fecfactu As String) As Boolean
Dim Letra As String, numasien As String
Dim cControlFra As CControlFacturaContab
Dim EstaEnTesoreria As String

    On Error GoTo EContab
    
    'Cojo la letra de serie
    Letra = Text1(0).Text
    
    Set cControlFra = New CControlFacturaContab
    numasien = ""
    
    'Con estos dos NO dejo pasar
    BuscaChekc = cControlFra.FechaCorrectaContabilizazion(ConnConta, CDate(fecfactu))
    If BuscaChekc <> "" Then numasien = numasien & "- " & BuscaChekc & vbCrLf
    BuscaChekc = cControlFra.FechaCorrectaIVA(ConnConta, CDate(fecfactu))
    If BuscaChekc <> "" Then numasien = numasien & "- " & BuscaChekc & vbCrLf
    Set cControlFra = Nothing
    
    If numasien <> "" Then
        FactContabilizada2 = True
        MsgBox numasien, vbExclamation
        Exit Function
    End If
    numasien = ""

    
    'Primero comprobaremos que esta el cobro en contabilidad
    EstaEnTesoreria = ""
    If Not ComprobarCobroArimoney(EstaEnTesoreria, Letra, CLng(numfactu), CDate(fecfactu)) Then
        FactContabilizada2 = True
        Exit Function
    End If

    'comprabar que se puede modificar/eliminar la factura
    If Me.Check1(1).Value = 1 Then 'si esta contabilizada
        'comprobar en la contabilidad si esta contabilizada
      
        If Letra <> "" Then
            If vParamAplic.ContabilidadNueva Then
                'Aunque en la nueva contabiliad SIEMPRE esta con apunte.
                numasien = DevuelveDesdeBDNew(cConta, "factcli", "numasien", "numserie", Letra, "T", , "numfactu", numfactu, "N", "anofactu", Year(fecfactu), "N")
            Else
                numasien = DevuelveDesdeBDNew(cConta, "cabfact", "numasien", "numserie", Letra, "T", , "codfaccl", numfactu, "N", "anofaccl", Year(fecfactu), "N")
            End If
            If Val(ComprobarCero(numasien)) <> 0 Then
'                FactContabilizada = True
'                MsgBox "La factura esta contabilizada y no se puede modificar.", vbInformation
'                Exit Function
            Else
                numasien = ""
            End If
        Else
'            MsgBox "Las factura de venta no tienen asignada una letra de serie", vbInformation
            numasien = ""
        End If
        
        Letra = "La factura esta en la contabilidad"
        If numasien <> "" Then Letra = Letra & vbCrLf & "Nº asiento: " & numasien
        Letra = Letra & vbCrLf & vbCrLf & "¿Continuar?"
        
        numasien = String(50, "*") & vbCrLf
        numasien = numasien & numasien & vbCrLf & vbCrLf
        Letra = numasien & Letra & vbCrLf & vbCrLf & numasien
        If MsgBox(Letra, vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
            FactContabilizada2 = False
        Else
            FactContabilizada2 = True
        End If
    Else
        FactContabilizada2 = False
    End If
    
    
EContab:
    If Err.Number <> 0 Then MuestraError Err.Number, "Comprobar Factura contabilizada", Err.Description
End Function


'En vTesoreria pondremos como estaba el recibo
'Es decir. El  msgbox que pondra al final lo guardo en esta variable
Private Function ComprobarCobroArimoney(vTesoreria As String, Letra As String, Codfaccl As Long, Fecha As Date) As Boolean
Dim vR As ADODB.Recordset
Dim Cad As String


On Error GoTo EComprobarCobroArimoney
    ComprobarCobroArimoney = False
    Set vR = New ADODB.Recordset
    
    If vParamAplic.ContabilidadNueva Then
        Cad = "Select * from cobros WHERE numserie='" & Letra & "'"
        Cad = Cad & " AND numfactu =" & Codfaccl
        Cad = Cad & " AND fecfactu =" & DBSet(Fecha, "F")
    Else
        Cad = "Select * from scobro WHERE numserie='" & Letra & "'"
        Cad = Cad & " AND codfaccl =" & Codfaccl
        Cad = Cad & " AND fecfaccl =" & DBSet(Fecha, "F")
    
    End If
    

    '
    vTesoreria = ""
    vR.Open Cad, ConnConta, adOpenForwardOnly, adLockPessimistic, adCmdText
    Cad = ""
    If vR.EOF Then
        vTesoreria = "NO se ha encotrado ningun vencimiento en la tesoreria"
    Else
        While Not vR.EOF
            Cad = ""
            If DBLet(vR!codrem, "T") <> "" Then
                Cad = "El cobro asociado a la factura esta remesado(" & vR!codrem & ")"
            Else
                If DBLet(vR!recedocu, "N") = 1 Then
                    Cad = "Documento recibido"
                Else
                    
                        If DBLet(vR!transfer, "N") = 1 Then
                            Cad = "Esta en una transferencia"
                        Else
                           If DBLet(vR!impcobro, "N") > 0 Then Cad = "Esta parcialmente cobrado: " & vR!impcobro
                        
                            
                                    'Si hubeira que poner mas coas iria aqui
                        End If 'transfer
                    
                End If 'recdedocu
            End If 'remesado
            If Cad <> "" Then vTesoreria = vTesoreria & "Vto: " & vR!numorden & "      " & Cad & vbCrLf
            vR.MoveNext
        Wend
    End If
    vR.Close
    
    
    
    If vTesoreria <> "" Then
        Cad = vTesoreria & vbCrLf & vbCrLf
        If vSesion.Nivel > 1 Then
            MsgBox Cad, vbExclamation
        Else
            Cad = Cad & "¿Continuar?"
            If MsgBox(Cad, vbQuestion + vbYesNo) = vbYes Then ComprobarCobroArimoney = True
        End If
    Else
        ComprobarCobroArimoney = True
    End If
            
EComprobarCobroArimoney:
    If Err.Number <> 0 Then MuestraError Err.Number
    Set vR = Nothing
End Function


