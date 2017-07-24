VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmComFacturar 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Facturas Compra Proveedores"
   ClientHeight    =   6315
   ClientLeft      =   45
   ClientTop       =   4335
   ClientWidth     =   11745
   Icon            =   "frmComFacturar.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmComFacturar.frx":000C
   ScaleHeight     =   6315
   ScaleWidth      =   11745
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameIntro 
      Height          =   1550
      Left            =   135
      TabIndex        =   8
      Top             =   495
      Width           =   11565
      Begin VB.CheckBox Check1 
         Caption         =   "Contabiliz."
         Height          =   375
         Index           =   1
         Left            =   5280
         TabIndex        =   50
         Top             =   1080
         Visible         =   0   'False
         Width           =   1240
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Tesoreria"
         Height          =   375
         Index           =   0
         Left            =   5295
         TabIndex        =   49
         Top             =   720
         Value           =   1  'Checked
         Visible         =   0   'False
         Width           =   1240
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   3
         Left            =   1400
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   15
         Text            =   "Text2"
         Top             =   1000
         Width           =   3735
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   2
         Left            =   3830
         MaxLength       =   10
         TabIndex        =   2
         Tag             =   "Fecha Recepción|F|N|||scafpc|fecrecep|dd/mm/yyyy|N|"
         Top             =   400
         Width           =   1305
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   5
         Left            =   7635
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   13
         Text            =   "Text2"
         Top             =   1000
         Width           =   3615
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   5
         Left            =   6915
         MaxLength       =   5
         TabIndex        =   4
         Text            =   "Text1"
         Top             =   1000
         Width           =   660
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   3
         Left            =   550
         MaxLength       =   6
         TabIndex        =   3
         Tag             =   "Cod. Proveedor|N|N|0|999999|scafpc|codprove|000000|S|"
         Text            =   "Text1"
         Top             =   1000
         Width           =   780
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   1
         Left            =   2005
         MaxLength       =   10
         TabIndex        =   1
         Tag             =   "Fecha Factura|F|N|||scafpc|fecfactu|dd/mm/yyyy|S|"
         Top             =   400
         Width           =   1305
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   0
         Left            =   240
         MaxLength       =   10
         TabIndex        =   0
         Tag             =   "Nº Factura|T|N|||scafpc|numfactu||S|"
         Text            =   "Text1 7"
         Top             =   400
         Width           =   1245
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   1
         Left            =   4895
         Picture         =   "frmComFacturar.frx":0A0E
         ToolTipText     =   "Buscar fecha"
         Top             =   150
         Width           =   240
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   0
         Left            =   3060
         Picture         =   "frmComFacturar.frx":0A99
         ToolTipText     =   "Buscar fecha"
         Top             =   135
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   2
         Left            =   6600
         ToolTipText     =   "Buscar banco propio"
         Top             =   1035
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   0
         Left            =   240
         ToolTipText     =   "Buscar proveedor"
         Top             =   1030
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha Recep."
         Height          =   255
         Index           =   3
         Left            =   3830
         TabIndex        =   14
         Top             =   200
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Cta. Prev. Pago"
         Height          =   255
         Index           =   2
         Left            =   6600
         TabIndex        =   12
         Top             =   795
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Proveedor"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   11
         Top             =   795
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha Factura"
         Height          =   255
         Index           =   29
         Left            =   2005
         TabIndex        =   10
         Top             =   200
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Nº Factura"
         Height          =   255
         Index           =   28
         Left            =   240
         TabIndex        =   9
         Top             =   200
         Width           =   1095
      End
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   4
      Left            =   2160
      MaxLength       =   4
      TabIndex        =   52
      Text            =   "Text1"
      Top             =   1110
      Width           =   660
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H80000018&
      Height          =   315
      Index           =   4
      Left            =   2880
      Locked          =   -1  'True
      MaxLength       =   40
      TabIndex        =   51
      Text            =   "Text2"
      Top             =   1110
      Width           =   3615
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   4050
      Left            =   120
      TabIndex        =   5
      Top             =   2080
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   7144
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.Frame FrameFactura 
      Height          =   4150
      Left            =   6840
      TabIndex        =   16
      Top             =   2000
      Width           =   4845
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   23
         Left            =   1440
         MaxLength       =   10
         TabIndex        =   54
         Text            =   "Text1"
         Top             =   3270
         Width           =   1380
      End
      Begin VB.CommandButton cmdGenerar 
         Caption         =   "&Generar"
         Height          =   375
         Left            =   1260
         TabIndex        =   36
         Top             =   3660
         Visible         =   0   'False
         Width           =   1035
      End
      Begin VB.CommandButton cmdCancelar 
         Cancel          =   -1  'True
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   90
         TabIndex        =   37
         Top             =   3660
         Visible         =   0   'False
         Width           =   1035
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Index           =   9
         Left            =   3000
         MaxLength       =   15
         TabIndex        =   43
         Tag             =   "Importe IVA 1|N|N|0||scafac|imporiv1|#,###,###,##0.00|N|"
         Text            =   "Text1 7"
         Top             =   1350
         Width           =   1485
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   8
         Left            =   3000
         MaxLength       =   15
         TabIndex        =   42
         Tag             =   "Base Imponible 3|N|N|0||scafac|baseimp3|#,###,###,##0.00|N|"
         Text            =   "Text1 7"
         Top             =   900
         Width           =   1485
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   7
         Left            =   3000
         MaxLength       =   15
         TabIndex        =   41
         Tag             =   "Base Imponible 2 |N|N|0||scafac|baseimp2|#,###,###,##0.00|N|"
         Text            =   "Text1 7"
         Top             =   570
         Width           =   1485
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   6
         Left            =   3000
         MaxLength       =   15
         TabIndex        =   39
         Tag             =   "Base Imponible 1|N|N|0||scafac|baseimp1|#,###,###,##0.00|N|"
         Text            =   "Text1 7"
         Top             =   240
         Width           =   1485
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   12
         Left            =   390
         MaxLength       =   5
         TabIndex        =   35
         Tag             =   "Codigo IVA 3|N|S|0|99|scafac|codiva3|00|N|"
         Text            =   "Text1 7"
         Top             =   2925
         Width           =   500
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   11
         Left            =   390
         MaxLength       =   5
         TabIndex        =   34
         Tag             =   "Codigo IVA 2|N|S|0|99|scafac|codiva2|00|N|"
         Text            =   "Text1 7"
         Top             =   2595
         Width           =   500
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   10
         Left            =   390
         MaxLength       =   5
         TabIndex        =   33
         Tag             =   "Codigo IVA 1|N|S|0|99|scafac|codiva1|00|N|"
         Text            =   "Text1 7"
         Top             =   2280
         Width           =   465
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   16
         Left            =   1440
         MaxLength       =   15
         TabIndex        =   26
         Tag             =   "Base Imponible 1|N|N|0||scafac|baseimp1|#,###,###,##0.00|N|"
         Text            =   "Text1 7"
         Top             =   2280
         Width           =   1395
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   13
         Left            =   885
         MaxLength       =   5
         TabIndex        =   25
         Tag             =   "% IVA 1|N|S|0|99.90|scafac|porciva1|#0.00|N|"
         Text            =   "Text1 7"
         Top             =   2280
         Width           =   525
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Index           =   19
         Left            =   3000
         MaxLength       =   15
         TabIndex        =   24
         Tag             =   "Importe IVA 1|N|N|0||scafac|imporiv1|#,###,###,##0.00|N|"
         Text            =   "Text1 7"
         Top             =   2280
         Width           =   1485
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   17
         Left            =   1440
         MaxLength       =   15
         TabIndex        =   23
         Tag             =   "Base Imponible 2 |N|N|0||scafac|baseimp2|#,###,###,##0.00|N|"
         Text            =   "Text1 7"
         Top             =   2595
         Width           =   1395
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   14
         Left            =   885
         MaxLength       =   5
         TabIndex        =   22
         Tag             =   "& IVA 2|N|S|0|99.90|scafac|porciva2|#0.00|N|"
         Text            =   "Text1 7"
         Top             =   2595
         Width           =   525
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Index           =   20
         Left            =   3000
         MaxLength       =   15
         TabIndex        =   21
         Tag             =   "Importe IVA 2|N|N|0||scafac|imporiv2|#,###,###,##0.00|N|"
         Text            =   "Text1 7"
         Top             =   2595
         Width           =   1485
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   18
         Left            =   1440
         MaxLength       =   15
         TabIndex        =   20
         Tag             =   "Base Imponible 3|N|N|0||scafac|baseimp3|#,###,###,##0.00|N|"
         Text            =   "Text1 7"
         Top             =   2925
         Width           =   1395
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   15
         Left            =   900
         MaxLength       =   5
         TabIndex        =   19
         Tag             =   "% IVA 3|N|S|0|99.90|scafac|porciva3|#0.00|N|"
         Text            =   "Text1 7"
         Top             =   2925
         Width           =   525
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Index           =   21
         Left            =   3000
         MaxLength       =   15
         TabIndex        =   18
         Tag             =   "Importe IVA 3|N|N|0||scafac|imporiv3|#,###,###,##0.00|N|"
         Text            =   "Text1 7"
         Top             =   2925
         Width           =   1485
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
         Height          =   350
         Index           =   22
         Left            =   2490
         MaxLength       =   15
         TabIndex        =   17
         Tag             =   "Total Factura|N|N|0||scafac|totalfac|#,###,###,##0.00|N|"
         Text            =   "Text1 7"
         Top             =   3720
         Width           =   1995
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha Vto"
         Height          =   195
         Index           =   12
         Left            =   270
         TabIndex        =   55
         Top             =   3330
         Width           =   825
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   2
         Left            =   1170
         Picture         =   "frmComFacturar.frx":0B24
         ToolTipText     =   "Buscar fecha"
         Top             =   3300
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   5
         Left            =   90
         ToolTipText     =   "Buscar codigo iva"
         Top             =   2940
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   4
         Left            =   90
         ToolTipText     =   "Buscar codigo iva"
         Top             =   2610
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   3
         Left            =   90
         ToolTipText     =   "Buscar codigo iva"
         Top             =   2310
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   11
         Left            =   2760
         TabIndex        =   48
         Top             =   900
         Width           =   135
      End
      Begin VB.Line Line2 
         X1              =   2520
         X2              =   4550
         Y1              =   1250
         Y2              =   1250
      End
      Begin VB.Label Label1 
         Caption         =   "Imp. dto. gnral."
         Height          =   255
         Index           =   10
         Left            =   1440
         TabIndex        =   47
         Top             =   900
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   9
         Left            =   2760
         TabIndex        =   46
         Top             =   570
         Width           =   135
      End
      Begin VB.Label Label1 
         Caption         =   "Imp. dto. ppago"
         Height          =   255
         Index           =   8
         Left            =   1440
         TabIndex        =   45
         Top             =   570
         Width           =   1215
      End
      Begin VB.Line Line1 
         X1              =   120
         X2              =   2880
         Y1              =   1920
         Y2              =   1920
      End
      Begin VB.Label Label1 
         Caption         =   "Base Imponible"
         Height          =   255
         Index           =   7
         Left            =   1440
         TabIndex        =   44
         Top             =   1350
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Bruto Factura"
         Height          =   255
         Index           =   6
         Left            =   1440
         TabIndex        =   40
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Cod."
         Height          =   255
         Index           =   5
         Left            =   390
         TabIndex        =   38
         Top             =   2070
         Width           =   375
      End
      Begin VB.Label Label1 
         Caption         =   "Base Imponible"
         Height          =   255
         Index           =   4
         Left            =   1440
         TabIndex        =   32
         Top             =   2040
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Importe IVA"
         Height          =   255
         Index           =   33
         Left            =   3000
         TabIndex        =   31
         Top             =   2070
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   37
         Left            =   3600
         TabIndex        =   30
         Top             =   1680
         Width           =   135
      End
      Begin VB.Label Label1 
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   36
         Left            =   11880
         TabIndex        =   29
         Top             =   2160
         Width           =   135
      End
      Begin VB.Label Label1 
         Caption         =   "TOTAL FACTURA"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Index           =   39
         Left            =   2970
         TabIndex        =   28
         Top             =   3450
         Width           =   1530
      End
      Begin VB.Label Label1 
         Caption         =   "% IVA"
         Height          =   255
         Index           =   41
         Left            =   885
         TabIndex        =   27
         Top             =   2070
         Width           =   495
      End
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   330
      Left            =   7080
      Top             =   5640
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
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
      Height          =   420
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   11745
      _ExtentX        =   20717
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   6
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Pedir Datos"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Ver Albaranes"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Generar Factura"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
         EndProperty
      EndProperty
      Begin VB.CheckBox chkVistaPrevia 
         Caption         =   "Vista previa"
         Height          =   195
         Left            =   8400
         TabIndex        =   7
         Top             =   120
         Visible         =   0   'False
         Width           =   1215
      End
   End
   Begin VB.Label Label1 
      Caption         =   "Operador"
      Height          =   255
      Index           =   1
      Left            =   1845
      TabIndex        =   53
      Top             =   900
      Width           =   735
   End
   Begin VB.Image imgBuscar 
      Height          =   240
      Index           =   1
      Left            =   1845
      Picture         =   "frmComFacturar.frx":0BAF
      ToolTipText     =   "Buscar trabajador"
      Top             =   1125
      Width           =   240
   End
   Begin VB.Menu mnOpciones 
      Caption         =   "&Opciones"
      Begin VB.Menu mnPedirDatos 
         Caption         =   "&Pedir Datos"
         Shortcut        =   ^D
      End
      Begin VB.Menu mnVerAlbaran 
         Caption         =   "&Ver Albaranes"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnGenerarFac 
         Caption         =   "&Generar Factura"
      End
      Begin VB.Menu mnBarra1 
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
Attribute VB_Name = "frmComFacturar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'========== VBLES PUBLICAS ====================
Public DatosADevolverBusqueda As String    'Tendra el nº de text que quiere que devuelva, empipados
Public Event DatoSeleccionado(CadenaSeleccion As String)


'========== VBLES PRIVADAS ====================
Private WithEvents frmF As frmCal 'Form Calendario Fecha
Attribute frmF.VB_VarHelpID = -1

Private WithEvents frmProv As frmManProve
Attribute frmProv.VB_VarHelpID = -1
'Private WithEvents frmT As frmAdmTrabajadores  'Form Mto Trabajadores
Private WithEvents frmBanPr As frmManBanco 'Mto de Bancos propios
Attribute frmBanPr.VB_VarHelpID = -1
Private WithEvents frmTipIVA As frmTipIVAConta  'Tipos de IVA de la contabilidad
Attribute frmTipIVA.VB_VarHelpID = -1
Private WithEvents frmFP As frmManFpago 'Form Mto Formas de Pago
Attribute frmFP.VB_VarHelpID = -1


Private Modo As Byte
'-----------------------------
'Se distinguen varios modos
'   0.-  Formulario limpio sin nungun campo rellenado
'   1.-  Preparando para hacer la busqueda
'   2.-  Ya tenemos registros y los vamos a recorrer
'        y podemos editarlos Edicion del campo
'   3.-  Insercion de nuevo registro
'   4.-  Modificar
'   5.-  Mantenimiento Lineas
'-------------------------------------------------------------------------
'-------------------------------------------------------------------------

'cadena donde se almacena la WHERE para la seleccion de los albaranes
'marcados para facturar
Dim cadWhere As String

'Cuando vuelve del formulario de ver los albaranes seleccionados hay que volver
'a cargar los datos de los albaranes
Dim VerAlbaranes As Boolean

Dim PrimeraVez As Boolean

'SQL de la tabla principal del formulario
Private CadenaConsulta As String

Private Ordenacion As String
Private NombreTabla As String  'Nombre de la tabla o de la
Private NomTablaLineas As String 'Nombre de la Tabla de lineas
Private kCampo As Integer
'-------------------------------------------------------------------------

Dim dtoGn As Currency
Dim dtoPP As Currency
Dim Forpa As Integer
Dim indCodigo As Integer

Dim BuscaChekc As String

Private vProve As CProveedor


Private Sub cmdCancelar_Click()
Dim I As Integer

    FrameIntro.Enabled = False
    ListView1.Enabled = True
    FrameFactura.Enabled = False
    
    
    BloquearTxt Text1(10), True
    BloquearTxt Text1(11), True
    BloquearTxt Text1(12), True
    
    
    For I = 3 To 5
        imgBuscar(I).Enabled = False
        imgBuscar(I).visible = False
    Next I
    
    
    Me.cmdCancelar.Enabled = False
    Me.cmdCancelar.visible = False
    Me.cmdGenerar.Enabled = False
    Me.cmdGenerar.visible = False

End Sub

Private Sub cmdGenerar_Click()
Dim I As Integer

    FrameIntro.Enabled = False
    ListView1.Enabled = True
    FrameFactura.Enabled = False


    For I = 3 To 5
        imgBuscar(I).Enabled = False
        imgBuscar(I).visible = False
    Next I
    
    
    Me.cmdCancelar.Enabled = False
    Me.cmdCancelar.visible = False
    Me.cmdGenerar.Enabled = False
    Me.cmdGenerar.visible = False


    BotonFacturar
    Set vProve = Nothing
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Activate()
    Screen.MousePointer = vbDefault
    If VerAlbaranes Then RefrescarAlbaranes
    VerAlbaranes = False
End Sub


Private Sub Form_Load()
Dim I As Integer

'    'Icono del formulario
'    Me.Icon = frmPpal.Icon
    
    ' ICONITOS DE LA BARRA
    With Me.Toolbar1
        .ImageList = frmPpal.imgListComun
        .Buttons(1).Image = 1   'Pedir Datos
        .Buttons(2).Image = 3   'Ver albaranes
        .Buttons(3).Image = 15   'Generar FActura
        .Buttons(6).Image = 11   'Salir
    End With
    
    'cargar IMAGES de busqueda
    Me.imgBuscar(0).Picture = frmPpal.imgListImages16.ListImages(1).Picture
    For I = 2 To 5
        Me.imgBuscar(I).Picture = frmPpal.imgListImages16.ListImages(1).Picture
    Next I
    
    Me.FrameFactura.Enabled = False
    For I = 3 To 5
        Me.imgBuscar(I).visible = False
        Me.imgBuscar(I).Enabled = False
    Next I
    
    '[Monica]28/01/2013: Si es Ribarroja pido la fecha de vencimiento
    FechaVtoVisible vParamAplic.Cooperativa = 5
    
    LimpiarCampos   'Limpia los campos TextBox
    InicializarListView
   
    '## A mano
    NombreTabla = "scafpc" 'cabecera facturas compras a proveedor
    NomTablaLineas = "slifpc" 'Tabla lineas de Facturacion
    Ordenacion = " ORDER BY scafpc.codprove, scafpc.numfactu, scafpc.fecfactu "
    
    'Vemos como esta guardado el valor del check
'    chkVistaPrevia.Value = CheckValueLeer(Name)
    
    'ASignamos un SQL al DATA1
    CadenaConsulta = "Select * from " & NombreTabla
    CadenaConsulta = CadenaConsulta & " where numfactu=-1"
    
    Data1.ConnectionString = Conn
    Data1.RecordSource = CadenaConsulta
    Data1.Refresh
        
    'Cargar el DataGrid de lineas de Revisiones inicialmente a nada DATA2
    PrimeraVez = True
    
    If DatosADevolverBusqueda = "" Then
        PonerModo 0
'    Else
'        PonerModo 1
    End If
    PrimeraVez = False
End Sub


Private Sub LimpiarCampos()
On Error Resume Next
    limpiar Me
    'Aqui va el especifico de cada form es
    '### a mano
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub Form_Unload(Cancel As Integer)
    CheckValueGuardar Me.Name, Me.chkVistaPrevia.Value
    DesBloqueoManual "RECFAC"
    TerminaBloquear
'    DesBloqueoManual ("scaalp")
End Sub


Private Sub frmBanPr_DatoSeleccionado(CadenaSeleccion As String)
    'Form de Mantenimiento de Bancos Propios
    Text1(5).Text = RecuperaValor(CadenaSeleccion, 1)
    Text1(5).Text = Format(Text1(5).Text, "0000")
    Text2(5).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub


Private Sub frmF_Selec(vFecha As Date)
Dim indice As Byte
    indice = CByte(Me.imgFecha(0).Tag)
    Text1(indice).Text = Format(vFecha, "dd/mm/yyyy")
End Sub

Private Sub frmFP_DatoSeleccionado(CadenaSeleccion As String)
'Form Mantenimiento de Formas de Pago
Dim indice As Byte
    
    Text1(indCodigo).Text = RecuperaValor(CadenaSeleccion, 1) 'Cod Forma Pago
    FormateaCampo Text1(indice)
    Text2(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2) 'Nom Forma Pago
End Sub


Private Sub frmProv_DatoSeleccionado(CadenaSeleccion As String)
'Mantenimiento de Proveedores
Dim indice As Byte
    
    indice = 3
    Text1(indice).Text = RecuperaValor(CadenaSeleccion, 1) 'Cod Proveedor
    FormateaCampo Text1(indice)
    Text2(indice).Text = RecuperaValor(CadenaSeleccion, 2) 'Nom proveedor
End Sub

Private Sub frmTipIVA_DatoSeleccionado(CadenaSeleccion As String)
'Tipos de IVA (de la Contabilidad)
    Text1(indCodigo).Text = RecuperaValor(CadenaSeleccion, 1) 'codigiva
    FormateaCampo Text1(indCodigo)
    Text1(indCodigo + 3).Text = RecuperaValor(CadenaSeleccion, 3) '% iva
    RecalcularDatosFactura
End Sub

'Private Sub frmT_DatoSeleccionado(CadenaSeleccion As String)
''Form Mantenimiento de Trabajadores
'Dim Indice As Byte
'    Indice = 4
'    Text1(Indice).Text = RecuperaValor(CadenaSeleccion, 1) 'Cod Trabajador
'    FormateaCampo Text1(Indice)
'    Text2(Indice).Text = RecuperaValor(CadenaSeleccion, 2) 'Nom Trabajador
'End Sub


Private Sub imgBuscar_Click(Index As Integer)
Dim indice As Byte

    If Modo = 2 Or Modo = 0 Then Exit Sub
    Screen.MousePointer = vbHourglass

    Select Case Index
        Case 0 'Cod. Proveedor
            Set frmProv = New frmManProve
            frmProv.DatosADevolverBusqueda = "0|"
            frmProv.Show vbModal
            Set frmProv = Nothing
            indice = 3
'--monica
'        Case 1 'Operador. Trabajador
'            Indice = 4
'            Set frmT = New frmAdmTrabajadores
'            frmT.DatosADevolverBusqueda = "0"
'            frmT.Show vbModal
'            Set frmT = Nothing
       
       Case 2 'Bancos Propios
            indice = 5
            Set frmBanPr = New frmManBanco
            frmBanPr.DatosADevolverBusqueda = "0|1|"
            frmBanPr.Show vbModal
            Set frmBanPr = Nothing
    
        Case 3, 4, 5 ' codigos de iva de contabilidad
            indCodigo = Index + 7
        
            Set frmTipIVA = New frmTipIVAConta
'            frmTipIva.DeConsulta = True
            frmTipIVA.DatosADevolverBusqueda = "0|1|2|"
            frmTipIVA.CodigoActual = Text1(indCodigo).Text
            frmTipIVA.Show vbModal
            Set frmTipIVA = Nothing
        
        Case 6 ' codigo de forma de pago
            indCodigo = 23
            
            Set frmFP = New frmManFpago
            frmFP.DatosADevolverBusqueda = "0|"
            frmFP.Show vbModal
            Set frmFP = Nothing
            
            indice = 23
    
    End Select
    
    PonerFoco Text1(indice)
    Screen.MousePointer = vbDefault
End Sub


Private Sub imgFecha_Click(Index As Integer)
Dim indice As Byte
Dim esq As Long
Dim dalt As Long
Dim menu As Long
Dim obj As Object

   If Modo = 2 Or Modo = 0 Then Exit Sub
   Screen.MousePointer = vbHourglass
    
   Set frmF = New frmCal
    
   esq = imgFecha(Index).Left
   dalt = imgFecha(Index).Top
    
   Set obj = imgFecha(Index).Container

   While imgFecha(Index).Parent.Name <> obj.Name
       esq = esq + obj.Left
       dalt = dalt + obj.Top
       Set obj = obj.Container
   Wend
    
   menu = Me.Height - Me.ScaleHeight 'ací tinc el heigth del menú i de la toolbar

   frmF.Left = esq + imgFecha(Index).Parent.Left + 30
   frmF.Top = dalt + imgFecha(Index).Parent.Top + imgFecha(Index).Height + menu - 40
   
   frmF.NovaData = Now
   indice = Index + 1
   
   If Index = 2 Then indice = 23
   
   Me.imgFecha(0).Tag = indice
   
   PonerFormatoFecha Text1(indice)
   If Text1(indice).Text <> "" Then frmF.NovaData = CDate(Text1(indice).Text)

   Screen.MousePointer = vbDefault
   frmF.Show vbModal
   Set frmF = Nothing
   PonerFoco Text1(indice)
End Sub

Private Sub ListView1_ItemCheck(ByVal item As MSComctlLib.ListItem)
'Cuando se selecciona un albaran de la lista
Dim I As Integer
Dim Cad As String
Dim TipoFP As Integer 'Forma de pago
Dim TipoDtoPP As Currency 'descuento pronto pago
Dim tipoDtoGn As Currency 'descuento general

    Screen.MousePointer = vbHourglass
    
    Set ListView1.SelectedItem = item
    
    'Inicializamos a cero
    TipoFP = 0
    TipoDtoPP = 0
    tipoDtoGn = 0
    
    'cuando seleccionamos un check vemos si lo podemos seleccionar
    'ya que si ya habia algun albaran selecionado tendremos que comprobar
    'que son de la misma forpa, dtoppago y dtognral.
    'si esto no se cumple no se pueden agrupar en la misma factura
    For I = 1 To ListView1.ListItems.Count
        If item.Index <> I Then
            If ListView1.ListItems(I).Checked Then
                'ya habia otro albaran seleccionado
                TipoFP = ListView1.ListItems(I).SubItems(2)
                TipoDtoPP = CCur(ListView1.ListItems(I).SubItems(4))
                tipoDtoGn = CCur(ListView1.ListItems(I).SubItems(5))
                Exit For
            End If
        End If
    Next I
    
    If Not (TipoFP = 0 And TipoDtoPP = 0 And tipoDtoGn = 0) Then
    'si ya habia un albaran seleccionado, comprobar que es del mismo tipo
        If item.SubItems(2) <> TipoFP Or item.SubItems(4) <> TipoDtoPP Or item.SubItems(5) <> tipoDtoGn Then
            MsgBox "Se debe seleccionar albaranes de la misma Forma de Pago y Descuentos", vbExclamation
            ListView1.SelectedItem.Checked = False
            Screen.MousePointer = vbDefault
            ListView1.SetFocus
            Exit Sub
        End If
    Else
    End If
    
    ' Calculamos los datos de factura
    If Not VerAlbaranes Then CalcularDatosFactura
    Screen.MousePointer = vbDefault
End Sub


Private Sub mnGenerarFac_Click()
Dim I As Integer
Dim rsVenci As ADODB.Recordset
Dim Rs As ADODB.Recordset
Dim FecVenci As Date
Dim FecVenci1 As Date
Dim SQL As String
Dim Cad As String
Dim ForPago As Long

'    If Text1(16).Text = "" And Text1(17).Text = "" And Text1(18).Text = "" Then
'        If MsgBox("Se va a generar una factura a cero. ¿ Desea continuar ?", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
'    End If


    FrameIntro.Enabled = False
    ListView1.Enabled = False
    FrameFactura.Enabled = True
    
    For I = 6 To 22
        BloquearTxt Text1(I), True
    Next I
    
    For I = 6 To 22
        Text1(I).Enabled = False
    Next I
    
    If vParamAplic.Cooperativa = 5 Then
        Text1(23).Enabled = True
        imgFecha(2).Enabled = True
        FecVenci = Text1(2).Text
        ' calculo el valor por defecto de la fecha de vto segun la forma de pago seleccionada de los albaranes
        
        'Obtener el Nº de Vencimientos de la forma de pago
        Cad = "select distinct codforpa from scaalp "
        Cad = Cad & " WHERE " & Replace(cadWhere, "slialp", "scaalp")
        ForPago = DevuelveValor(Cad)
        
        
        SQL = "SELECT numerove, diasvto, restoven FROM sforpa WHERE codforpa=" & ForPago
        Set rsVenci = New ADODB.Recordset
        rsVenci.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
        If Not rsVenci.EOF Then
            If rsVenci!numerove > 0 Then
                'Obtener los dias de pago de la tabla de parametros: spara1
                SQL = " SELECT  diapago1, diapago2, diapago3,mesnogir "
                SQL = SQL & " FROM sparam "
                SQL = SQL & " WHERE codparam=1"
                Set Rs = New ADODB.Recordset
                Rs.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                
                If Not Rs.EOF Then
                  'FECHA VTO
                  FecVenci = CDate(Text1(2).Text)
                  '=== Modificado: Laura 23/01/2007
    '              FecVenci = FecVenci + CByte(DBLet(rsVenci!primerve, "N"))
                  FecVenci = DateAdd("d", DBLet(rsVenci!diasvto, "N"), FecVenci)
                  
                  'comprobar si tiene dias de pago y obtener la fecha del vencimiento correcta
                  FecVenci = ComprobarFechaVenci(FecVenci, DBLet(Rs!DiaPago1, "N"), DBLet(Rs!DiaPago2, "N"), DBLet(Rs!DiaPago3, "N"))
    
                  'Comprobar si  tiene mes a no girar
                  FecVenci1 = FecVenci
                  If DBSet(Rs!mesnogir, "N") <> 0 Then
                      FecVenci1 = ComprobarMesNoGira(FecVenci1, DBSet(Rs!mesnogir, "N"), DBSet(0, "N"), Rs!DiaPago1, Rs!DiaPago2, Rs!DiaPago3)
                  End If
                  
                  Text1(23).Text = FecVenci1
                  '==================================
                End If
            End If
        End If
    End If
    
    BloquearTxt Text1(10), (Text1(16).Text = "")
    BloquearTxt Text1(11), (Text1(17).Text = "")
    BloquearTxt Text1(12), (Text1(18).Text = "")
    
    imgBuscar(3).Enabled = (Text1(16).Text <> "")
    imgBuscar(4).Enabled = (Text1(17).Text <> "")
    imgBuscar(5).Enabled = (Text1(18).Text <> "")
    imgBuscar(3).visible = (Text1(16).Text <> "")
    imgBuscar(4).visible = (Text1(17).Text <> "")
    imgBuscar(5).visible = (Text1(18).Text <> "")
    
    Me.cmdCancelar.Enabled = True
    Me.cmdCancelar.visible = True
    Me.cmdGenerar.Enabled = True
    Me.cmdGenerar.visible = True
    
    PonerFoco Text1(10)
    
'    BotonFacturar
'    Set vProve = Nothing
'    Screen.MousePointer = vbDefault
End Sub

Private Sub mnPedirDatos_Click()
    BotonPedirDatos
End Sub

Private Sub mnSalir_Click()
    Screen.MousePointer = vbDefault
    Unload Me
End Sub


Private Sub mnVerAlbaran_Click()
    BotonVerAlbaranes
End Sub

Private Sub Text1_GotFocus(Index As Integer)
    kCampo = Index
    ConseguirFoco Text1(Index), Modo
End Sub


Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'Avanzar/Retroceder los campos con las flechas de desplazamiento del teclado.
    KEYdown KeyCode
End Sub


Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


'----------------------------------------------------------------
'----------------------------------------------------------------
' Cunado el campo de texto pierde el enfoque
' Es especifico de cada formulario y en el podremos controlar
' lo que queramos, desde formatear un campo si asi lo deseamos
' hasta pedir que nos devuelva los datos de la empresa
'----------------------------------------------------------------
'----------------------------------------------------------------
Private Sub Text1_LostFocus(Index As Integer)
        
    If Not PerderFocoGnral(Text1(Index), Modo) Then Exit Sub
        
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub
          
    'Si queremos hacer algo ..
    Select Case Index
        Case 1, 2 'Fecha factura, fecha recepcion
            PonerFormatoFecha Text1(Index)
            If Text1(Index) <> "" Then
                ' No debe existir el número de factura para el proveedor en hco
                If ExisteFacturaEnHco Then
                    InicializarListView
                End If
            End If
            
        Case 23 ' fecha de vencimiento
            PonerFormatoFecha Text1(Index)
            
        Case 3 'Cod Proveedor
            If PonerFormatoEntero(Text1(Index)) Then
                Text2(Index).Text = PonerNombreDeCod(Text1(Index), "proveedor", "nomprove", "codprove")
                
                Text1(5).Text = DevuelveDesdeBDNew(cPTours, "proveedor", "codbanpr", "codprove", Text1(3).Text, "N")
                If Text1(5).Text <> "" Then Text2(5).Text = DevuelveDesdeBDNew(cPTours, "sbanco", "nombanco", "codbanpr", Text1(5).Text, "N")
                
                ' No debe existir el número de factura para el proveedor en hco
                If ExisteFacturaEnHco Then
                    InicializarListView
                Else
                    'comprobamos que no haya nadie recepcionando facturas de ese proveedor
                    DesBloqueoManual ("RECFAC")
                    If Not BloqueoManual("RECFAC", Text1(3).Text) Then
                        MsgBox "No se puede recepcionar factura de ese proveedor. Hay otro usuario recepcionando.", vbExclamation
                        BotonPedirDatos
                        Screen.MousePointer = vbDefault
                        Exit Sub
                    Else
                        CargarAlbaranes
                    End If
                    
                End If
                
            Else
                Text2(Index).Text = ""
            End If
            
'--monica
'        Case 4 'Cod Trabajador
'            If PonerFormatoEntero(Text1(Index)) Then
'                Text2(Index).Text = PonerNombreDeCod(Text1(Index), conAri, "straba", "nomtraba", "codtraba")
'            Else
'                Text2(Index).Text = ""
'            End If

        Case 5 'Cta Prevista de PAgo
            If PonerFormatoEntero(Text1(Index)) Then
                Text2(Index).Text = PonerNombreDeCod(Text1(Index), "sbanco", "nombanco", "codbanpr")
                Text1(Index).Text = Format(Text1(Index).Text, "00")
            Else
                Text2(Index).Text = ""
            End If
            
        Case 10, 11, 12 ' codigo de iva
            If PonerFormatoEntero(Text1(Index)) Then
                Text1(Index + 3).Text = DevuelveDesdeBDNew(cConta, "tiposiva", "porceiva", "codigiva", Text1(Index).Text, "N")
            Else
                Text1(Index + 3).Text = ""
            End If
        
        
            RecalcularDatosFactura
            
    End Select
End Sub


'----------------------------------------------------------------
'----------------------------------------------------------------
'   En PONERMODO se habilitan, o no, los diverso campos del
'   formulario en funcion del modo en k vayamos a trabajar
Private Sub PonerModo(Kmodo As Byte)
Dim I As Byte, Numreg As Byte
Dim b As Boolean
On Error GoTo EPonerModo

    Modo = Kmodo
    
    'Modo 2. Hay datos y estamos visualizandolos
    '=========================================
    b = (Modo = 2)
        
                 
    'Bloquea los campos Text1 sino estamos modificando/Insertando Datos
    'Si estamos en Insertar además limpia los campos Text1
    'si estamos en modificar bloquea las compos que son clave primaria
    BloquearText1 Me, Modo
    
    'Importes siempre bloqueados
    For I = 6 To 22
        BloquearTxt Text1(I), True
    Next I
    'Campo B.Imp y Imp. IVA siempre en azul
    Text1(9).BackColor = &HFFFFC0 'Base imponible
    Text1(19).BackColor = &HFFFFC0 'Total Iva 1
    Text1(20).BackColor = &HFFFFC0 'Iva 2
    Text1(21).BackColor = &HFFFFC0 'IVa 3
    Text1(22).BackColor = &HC0C0FF    'Total factura
    
    
    '---------------------------------------------
    b = (Modo <> 0 And Modo <> 2 And Modo <> 5)
    
    For I = 0 To Me.imgBuscar.Count - 1
        Me.imgBuscar(I).Enabled = b
    Next I
                    
    Me.chkVistaPrevia.Enabled = (Modo <= 2)
       
    'Poner el tamaño de los campos. Si es modo Busqueda el MaxLength del campo
    'debe ser mayor para adminir intervalos de busqueda.
    PonerLongCampos
'    PonerModoOpcionesMenu (Modo) 'Activar opciones de menu según modo
    PonerOpcionesMenu 'Activar opciones de menu según nivel de permisos del usuario
    
EPonerModo:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub


Private Sub PonerLongCampos()
'Modificar el MaxLength del campo en funcion de si es modo de búsqueda o no
'para los campos que permitan introducir criterios más largos del tamaño del campo
    PonerLongCamposGnral Me, Modo, 1
End Sub


Private Function DatosOk() As Boolean
'Comprobar que los datos del frame de introduccion son correctos antes de cargar datos
Dim vtag As CTag
Dim Cad As String
Dim I As Byte
Dim cta As String
Dim cadMen As String


    On Error GoTo EDatosOK
    DatosOk = False
    
    ' deben de introducirse todos los datos del frame
    For I = 0 To 5
        If Text1(I).Text = "" And I <> 4 Then
            If Text1(I).Tag <> "" Then
                Set vtag = New CTag
                If vtag.Cargar(Text1(I)) Then
                    Cad = vtag.Nombre
                Else
                    Cad = "Campo"
                End If
                Set vtag = Nothing
            Else
                Cad = "Campo"
                If I = 5 Then Cad = "Cta. Prev. Pago"
            End If
            MsgBox Cad & " no puede estar vacio. Reintroduzca", vbExclamation
            PonerFoco Text1(I)
            Exit Function
        End If
    Next I
        
    'comprobar que la fecha de la factura sea anterior a la fecha de recepcion
    If Not EsFechaIgualPosterior(Text1(1).Text, Text1(2).Text, True, "La fecha de recepción debe ser igual o posterior a la fecha de la factura.") Then
        Exit Function
    End If
    
    'Comprobar que la fecha de RECEPCION esta dentro de los ejercicios contables
    If vParamAplic.NumeroConta <> 0 Then
        '[Monica]20/06/2017: como David
        ResultadoFechaContaOK = EsFechaOKConta(CDate(Text1(2).Text))
        If ResultadoFechaContaOK > 0 Then
            If ResultadoFechaContaOK <> 4 Then MsgBox MensajeFechaOkConta, vbExclamation
            Exit Function
        End If
        
        
        
        
        ' comprobamos que el proveedor tenga cuenta contable para introducir el pago
        If vProve.CuentaCble = "" Then
            MsgBox "El proveedor no tiene asignada una cuenta contable", vbExclamation
            Exit Function
        End If
        
    End If
    
    'comprobar que se han seleccionado lineas para facturar
    If cadWhere = "" Then
        MsgBox "Debe seleccionar albaranes para facturar.", vbExclamation
        Exit Function
    End If
    
    
    ' No debe existir el número de factura para el proveedor en hco
    If ExisteFacturaEnHco Then Exit Function
    
    
    'todos los albaranes seleccionados deben tener la misma: forma pago, dto ppago, dto gnral
    Cad = "select count(distinct codforpa,dtoppago,dtognral) from scaalp "
    Cad = Cad & " WHERE " & Replace(cadWhere, "slialp", "scaalp")
    If RegistrosAListar(Cad) > 1 Then
        MsgBox "No se puede facturar albaranes con distintas: forma de pago, dto gral, dto ppago.", vbExclamation
        Exit Function
    End If
    
    
    'Si la forpa es TRANSFERENCIA entonces compruebo la si tiene cta bancaria
    Cad = "select distinct (codforpa) from scaalp "
    Cad = Cad & " WHERE " & Replace(cadWhere, "slialp", "scaalp")
    Set miRsAux = New ADODB.Recordset
    miRsAux.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Cad = miRsAux.Fields(0)
    miRsAux.Close
    
    
    
    'Ahora buscamos el tipforpa del codforpa
    Cad = "Select tipforpa from sforpa where codforpa=" & Cad
    miRsAux.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    I = 0
    If miRsAux.EOF Then
        MsgBox "Error en el TIPO de forma de pago", vbExclamation
    Else
        I = 1
        Cad = miRsAux.Fields(0)
        If Val(Cad) = vbFPTransferencia Then
            'Compruebo que la forpa es transferencia
            I = 2
        End If
    End If
    miRsAux.Close
    Set miRsAux = Nothing
    
    
    If I = 2 Then
        'La forma de pago es transferencia. Debo comprobar que existe la cuenta bancaria
        'del proveedor
        If vProve.CuentaBan = "" Or vProve.DigControl = "" Or vProve.Sucursal = "" Or vProve.Banco = "" Then
            Cad = "Cuenta bancaria incorrecta. Forma de pago: transferencia.    ¿Continuar?"
            If MsgBox(Cad, vbQuestion + vbYesNoCancel) <> vbYes Then I = 0

'******
'[Monica]22/11/2013: Tema del iban
        Else
            cta = Format(vProve.Banco, "0000") & Format(vProve.Sucursal, "0000") & Format(vProve.DigControl, "00") & Format(vProve.CuentaBan, "0000000000")
            If Not Comprueba_CC(cta) Then
                cadMen = "La cuenta bancaria no es correcta. ¿ Desea continuar ?."
                If MsgBox(cadMen, vbQuestion + vbYesNo + vbDefaultButton1) = vbYes Then
                    I = 0
                End If
            Else
                BuscaChekc = ""
                If vProve.IBAN <> "" Then BuscaChekc = Mid(vProve.IBAN, 1, 2)
                    
                If DevuelveIBAN2(BuscaChekc, cta, cta) Then
                    If Mid(vProve.IBAN, 3) <> cta Then
                        cta = "Calculado : " & BuscaChekc & cta
                        cta = "Introducido: " & vProve.IBAN & vbCrLf & cta & vbCrLf
                        cta = "Error en codigo IBAN" & vbCrLf & cta & "Continuar?"
                        If MsgBox(cta, vbQuestion + vbYesNo) = vbYes Then I = 0
                    End If
                End If
            End If
'******
        End If
    End If
    
    'Si i=0 es que o esta mal la forpa o no quiere seguir pq no tiene cuenta bancaria
    If I > 0 Then DatosOk = True
    Exit Function
    
EDatosOK:
    DatosOk = False
    MuestraError Err.Number, "Comprobar datos correctos", Err.Description
End Function



Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1  'Pedir datos
             mnPedirDatos_Click
             
        Case 2 'Ver Albaranes
            mnVerAlbaran_Click
            
        Case 3 'Generar Factura
            mnGenerarFac_Click

        Case 6    'Salir
            mnSalir_Click
    End Select
End Sub


Private Sub PonerOpcionesMenu()
Dim J As Byte

    PonerOpcionesMenuGeneral Me
    
    J = Val(Me.mnPedirDatos.HelpContextID)
    If J < vSesion.Nivel Then Me.mnPedirDatos.Enabled = False
    
    J = Val(Me.mnGenerarFac.HelpContextID)
    If J < vSesion.Nivel Then Me.mnGenerarFac.Enabled = False
End Sub


Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, Modo, cerrar
    If cerrar Then Unload Me
End Sub

 
Private Sub BotonPedirDatos()
Dim Nombre As String


    FrameIntro.Enabled = True
    ListView1.Enabled = False
    FrameFactura.Enabled = False
    

    'Vaciamos todos los Text
    LimpiarCampos
    'Vaciamos el ListView
    InicializarListView
    
    'Como no habrá albaranes seleccionados vaciamos la cadwhere
    cadWhere = ""
    
    PonerModo 3
    
    'fecha recepcion
    Text1(2).Text = Format(Now, "dd/mm/yyyy")
    
'--monica
'    'poner trabajador conectado como operador
'    Text1(4).Text = PonerTrabajadorConectado(Nombre)
'    Text2(4).Text = Nombre
    
    'desbloquear los registros de la saalp (si hay bloquedos)
    TerminaBloquear
    
    'si vamos
    'desBloqueo Manual de las tablas
'    DesBloqueoManual ("scaalp")
    
    PonerFoco Text1(0)
End Sub

Private Sub BotonModificar()
Dim Nombre As String

    PonerModo 4
    
    PonerFoco Text1(10)

End Sub




Private Sub BotonVerAlbaranes()

    If Not SeleccionaRegistros Then Exit Sub
    
    VerAlbaranes = True
    
    frmComEntAlbaranes.cadSelAlbaranes = cadWhere
    frmComEntAlbaranes.EsHistorico = False
    frmComEntAlbaranes.Show vbModal
    frmComEntAlbaranes.cadSelAlbaranes = ""
End Sub
    


Private Sub CargarAlbaranes()
'Recupera de la BD y muestra en el Listview todos los albaranes de compra
'que tiene el proveedor introducido.
Dim SQL As String
Dim Rs As ADODB.Recordset
Dim ItmX As ListItem
On Error GoTo ECargar

    ListView1.ListItems.Clear
    If VerAlbaranes = False Then cadWhere = ""
    
    'si no hay proveedor salir
    If Text1(3).Text = "" Then Exit Sub
    
    SQL = "SELECT scaalp.numalbar,scaalp.fechaalb,scaalp.codforpa,sforpa.nomforpa,scaalp.dtoppago,scaalp.dtognral, "
    SQL = SQL & " sum(slialp.importel) as bruto "
    SQL = SQL & " FROM (scaalp LEFT OUTER JOIN sforpa ON scaalp.codforpa=sforpa.codforpa) "
    SQL = SQL & " INNER JOIN slialp ON scaalp.numalbar = slialp.numalbar  AND scaalp.fechaalb=slialp.fechaalb AND scaalp.codprove=slialp.codprove "
    SQL = SQL & " WHERE scaalp.codprove =" & Text1(3).Text
    SQL = SQL & " GROUP BY scaalp.numalbar, scaalp.fechaalb, scaalp.codforpa, scaalp.dtoppago,scaalp.dtognral "
    SQL = SQL & " ORDER BY scaalp.numalbar"

    Set Rs = New ADODB.Recordset
    Rs.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    
    InicializarListView
    
    While Not Rs.EOF
        Set ItmX = ListView1.ListItems.Add()
        ItmX.Text = Rs!numalbar
        ItmX.SubItems(1) = Format(Rs!fechaalb, "dd/mm/yyyy")
        ItmX.SubItems(2) = Format(Rs!Codforpa, "000")
        ItmX.SubItems(3) = Rs!nomforpa
        ItmX.SubItems(4) = Format(Rs!DtoPPago, "#0.00")
        ItmX.SubItems(5) = Format(Rs!DtoGnral, "#0.00")
        ItmX.SubItems(6) = Format(Rs!Bruto, "#,###,#0.00") '(RAFA/ALZIRA) 12092006
        'Sig
        Rs.MoveNext
    Wend
    Rs.Close
    Set Rs = Nothing
    
    ListView1.Enabled = True

ECargar:
    If Err.Number <> 0 Then MuestraError Err.Number, "Cargando Albaranes", Err.Description
End Sub


Private Sub InicializarListView()
'Inicializa las columnas del List view

    ListView1.ListItems.Clear
    
    ListView1.ColumnHeaders.Clear
    ListView1.ColumnHeaders.Add , , "NºAlbaran", 1100
    ListView1.ColumnHeaders.Add , , "Fecha", 1100, 2
    ListView1.ColumnHeaders.Add , , "FPag", 550
    ListView1.ColumnHeaders.Add , , "Desc. FPago", 1450
    ListView1.ColumnHeaders.Add , , "DtoPP", 650, 2
    ListView1.ColumnHeaders.Add , , "DtoGr", 600, 2
    ListView1.ColumnHeaders.Add , , "Imp. Bruto", 1100, 1
End Sub



Private Sub CalcularDatosFactura()
Dim I As Integer
Dim SQL As String
Dim cadAux As String
Dim vFactu As CFacturaCom

    'Limpiar en el form los datos calculados de la factura
    'y volvemos a recalcular
    For I = 6 To 22
         Text1(I).Text = ""
    Next I
    
    cadAux = ""
    cadWhere = ""
    
    For I = 1 To ListView1.ListItems.Count
        If ListView1.ListItems(I).Checked Then
        'para cada albaran seleccionado para la factura
            Forpa = ListView1.ListItems(I).SubItems(2)
            dtoPP = ListView1.ListItems(I).SubItems(4)
            dtoGn = ListView1.ListItems(I).SubItems(5)
            SQL = "(numalbar=" & DBSet(ListView1.ListItems(I).Text, "T") & " and "
            SQL = SQL & "fechaalb=" & DBSet(ListView1.ListItems(I).SubItems(1), "F") & ")"
            If cadAux = "" Then
                cadAux = SQL
            Else
                cadAux = cadAux & " OR " & SQL
            End If
        End If
    Next I
    
    If cadAux <> "" Then
    'se han seleccionado albaranes para facturar
    'Esta el la cadena WHERE de los albaranes seleccionados para obtener
    'el bruto de las lineas de los albaranes agrupadas por tipo de iva
        cadWhere = "slialp.codprove=" & Val(Text1(3).Text)
        cadWhere = cadWhere & " AND (" & cadAux & ")"
    Else
        Exit Sub
    End If
    
    
    If Not SeleccionaRegistros Then Exit Sub
    
    If Not BloqueaRegistro("scaalp", cadWhere) Then
        ListView1.SelectedItem.Checked = False
    End If
    
    Set vFactu = New CFacturaCom
    vFactu.DtoPPago = dtoPP
    vFactu.DtoGnral = dtoGn
    If vFactu.CalcularDatosFactura(cadWhere, "scaalp", "slialp") Then
        Text1(6).Text = vFactu.BrutoFac
        Text1(7).Text = vFactu.ImpPPago
        Text1(8).Text = vFactu.ImpGnral
        Text1(9).Text = vFactu.BaseImp
        Text1(10).Text = vFactu.TipoIVA1
        Text1(11).Text = vFactu.TipoIVA2
        Text1(12).Text = vFactu.TipoIVA3
        Text1(13).Text = vFactu.PorceIVA1
        Text1(14).Text = vFactu.PorceIVA2
        Text1(15).Text = vFactu.PorceIVA3
        Text1(16).Text = vFactu.BaseIVA1
        Text1(17).Text = vFactu.BaseIVA2
        Text1(18).Text = vFactu.BaseIVA3
        Text1(19).Text = vFactu.ImpIVA1
        Text1(20).Text = vFactu.ImpIVA2
        Text1(21).Text = vFactu.ImpIVA3
        Text1(22).Text = vFactu.TotalFac
        
        For I = 6 To 22
            FormateaCampo Text1(I)
        Next I
        'Quitar ceros de linea IVA 2
        If Val(Text1(14).Text) = 0 And Val(Text1(11).Text) = 0 Then
            For I = 11 To 20 Step 3
                Text1(I).Text = QuitarCero(CCur(Text1(I).Text))
            Next I
        End If
        'Quitar ceros de linea IVA 3
        If Val(Text1(15).Text) = 0 And Val(Text1(12).Text) = 0 Then
            For I = 12 To 21 Step 3
                Text1(I).Text = QuitarCero(CCur(Text1(I).Text))
            Next I
        End If
        
    Else
        MuestraError Err.Number, "Calculando Factura", Err.Description
    End If
    Set vFactu = Nothing
   
End Sub


Private Sub RecalcularDatosFactura()
Dim I As Integer
Dim SQL As String
Dim cadAux As String
Dim TotalFactura As Currency
    
Dim ImpBImIVA As Currency
Dim ImpIva As Currency
Dim ImpIVA1 As Currency
Dim ImpIVA2 As Currency
Dim ImpIVA3 As Currency
    
    On Error GoTo eRecalcularDatosFactura
    
    If Not SeleccionaRegistros Then Exit Sub
    
    If Not BloqueaRegistro("scaalp", cadWhere) Then
        ListView1.SelectedItem.Checked = False
    End If
    
    TotalFactura = 0
    If Text1(16).Text <> "" Then
        cadAux = Text1(13).Text
        ImpBImIVA = CCur(ImporteSinFormato(Text1(16).Text))
        If cadAux = "" Then cadAux = "0"
        ImpIva = CalcularPorcentaje(ImpBImIVA, CCur(cadAux), 2)
        
        ImpIVA1 = ImpIva
        
        
        'sumamos todos los IVAS para sumarselo a la base imponible total de la factura
        'los vamos acumulando
        TotalFactura = TotalFactura + ImpBImIVA + ImpIva
    End If
    
    If Text1(17).Text <> "" Then
        cadAux = Text1(14).Text
        ImpBImIVA = CCur(ImporteSinFormato(Text1(17).Text))
        If cadAux = "" Then cadAux = "0"
        ImpIva = CalcularPorcentaje(ImpBImIVA, CCur(cadAux), 2)
        
        ImpIVA2 = ImpIva
        
        
        'sumamos todos los IVAS para sumarselo a la base imponible total de la factura
        'los vamos acumulando
        TotalFactura = TotalFactura + ImpBImIVA + ImpIva
    End If
    
    
    If Text1(18).Text <> "" Then
        cadAux = Text1(15).Text
        ImpBImIVA = CCur(ImporteSinFormato(Text1(18).Text))
        If cadAux = "" Then cadAux = "0"
        ImpIva = CalcularPorcentaje(ImpBImIVA, CCur(cadAux), 2)
        
        ImpIVA3 = ImpIva
        
        
        'sumamos todos los IVAS para sumarselo a la base imponible total de la factura
        'los vamos acumulando
        TotalFactura = TotalFactura + ImpBImIVA + ImpIva
    End If
        
'        Text1(6).Text = vFactu.BrutoFac
'        Text1(7).Text = vFactu.ImpPPago
'        Text1(8).Text = vFactu.ImpGnral
'        Text1(9).Text = vFactu.BaseImp
'        Text1(10).Text = vFactu.TipoIVA1
'        Text1(11).Text = vFactu.TipoIVA2
'        Text1(12).Text = vFactu.TipoIVA3
'        Text1(13).Text = vFactu.PorceIVA1
'        Text1(14).Text = vFactu.PorceIVA2
'        Text1(15).Text = vFactu.PorceIVA3
'        Text1(16).Text = vFactu.BaseIVA1
'        Text1(17).Text = vFactu.BaseIVA2
'        Text1(18).Text = vFactu.BaseIVA3
        
        Text1(19).Text = ImpIVA1
        Text1(20).Text = ImpIVA2
        Text1(21).Text = ImpIVA3
        Text1(22).Text = TotalFactura
        
        For I = 19 To 22
            FormateaCampo Text1(I)
        Next I
        'Quitar ceros de linea IVA 2
        If Val(Text1(14).Text) = 0 And Val(Text1(11).Text) = 0 Then
            For I = 11 To 20 Step 3
                Text1(I).Text = ""
            Next I
        End If
        'Quitar ceros de linea IVA 3
        If Val(Text1(15).Text) = 0 And Val(Text1(12).Text) = 0 Then
            For I = 12 To 21 Step 3
                Text1(I).Text = ""
            Next I
        End If
        Exit Sub
        
   
eRecalcularDatosFactura:
    MuestraError Err.Number, "Recalculando Datos de Factura", Err.Description
End Sub





Private Function SeleccionaRegistros() As Boolean
'Comprueba que se seleccionan albaranes en la base de datos
'es decir que hay albaranes marcados
'cuando se van marcando albaranes se van añadiendo el la cadena cadWhere
Dim SQL As String

    On Error GoTo ESel
    SeleccionaRegistros = False
    
    If cadWhere = "" Then Exit Function
    cadWhere = Replace(cadWhere, "slialp", "scaalp")
    
    SQL = "Select count(*) FROM scaalp"
    SQL = SQL & " WHERE " & cadWhere
    If RegistrosAListar(SQL) <> 0 Then SeleccionaRegistros = True
    Exit Function
    
ESel:
    SeleccionaRegistros = False
    MuestraError Err.Number, "No hay seleccionados Albaranes", Err.Description
End Function


Private Sub BotonFacturar()
Dim vFactu As CFacturaCom
Dim Cad As String
Dim I As Integer


    Screen.MousePointer = vbHourglass
    On Error GoTo Error1
    
    ' preguntamos antes de recepcionar
    If MsgBox("¿ Desea generar la Factura de Proveedor ?", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
    
    Cad = ""
    If Text1(3).Text = "" Then
        Cad = "Falta proveedor"
    Else
        If Not IsNumeric(Text1(3).Text) Then Cad = "Campo proveedor debe ser numérico"
    End If
    If Cad <> "" Then
        MsgBox Cad, vbExclamation
        Exit Sub
    End If
        
        
        
    Set vProve = New CProveedor
    
    'Tiene que ller los datos del proveedor
    If Not vProve.LeerDatos(Text1(3).Text) Then Exit Sub
    
    
    If Not DatosOk Then
        Exit Sub
    End If
    
        'Pasar los Albaranes seleccionados con cadWHERE a una factura
        Set vFactu = New CFacturaCom
        vFactu.Proveedor = Text1(3).Text
        vFactu.numfactu = Text1(0).Text
        vFactu.Fecfactu = Text1(1).Text
        vFactu.FecRecep = Text1(2).Text
        vFactu.Trabajador = Text1(4).Text
        vFactu.BancoPr = Text1(5).Text
        vFactu.BrutoFac = ImporteFormateado(Text1(6).Text)
        vFactu.ForPago = Forpa
        vFactu.DtoPPago = dtoPP
        vFactu.DtoGnral = dtoGn
        vFactu.ImpPPago = ImporteFormateado(Text1(7).Text)
        vFactu.ImpGnral = ImporteFormateado(Text1(8).Text)
        vFactu.BaseIVA1 = ImporteFormateado(Text1(16).Text)
        vFactu.BaseIVA2 = ImporteFormateado(Text1(17).Text)
        vFactu.BaseIVA3 = ImporteFormateado(Text1(18).Text)
        vFactu.TipoIVA1 = ComprobarCero(Text1(10).Text)
        vFactu.TipoIVA2 = ComprobarCero(Text1(11).Text)
        vFactu.TipoIVA3 = ComprobarCero(Text1(12).Text)
        vFactu.PorceIVA1 = ComprobarCero(Text1(13).Text)
        vFactu.PorceIVA2 = ComprobarCero(Text1(14).Text)
        vFactu.PorceIVA3 = ComprobarCero(Text1(15).Text)
        vFactu.ImpIVA1 = ImporteFormateado(Text1(19).Text)
        vFactu.ImpIVA2 = ImporteFormateado(Text1(20).Text)
        vFactu.ImpIVA3 = ImporteFormateado(Text1(21).Text)
        vFactu.TotalFac = ImporteFormateado(Text1(22).Text)
        If vParamAplic.Cooperativa = 5 Then
            vFactu.FecVto = Text1(23).Text
        Else
            vFactu.FecVto = ""
        End If
        
        'Si el proveedor tiene CTA BANCARIA se la asigno
        vFactu.CCC_Entidad = vProve.Banco
        vFactu.CCC_Oficina = vProve.Sucursal
        vFactu.CCC_CC = vProve.DigControl
        vFactu.CCC_CTa = vProve.CuentaBan
        '[Monica]22/11/2013: Tema iban
        vFactu.CCC_Iban = vProve.IBAN
        
        If vFactu.TraspasoAlbaranesAFactura(cadWhere) Then BotonPedirDatos
        Set vFactu = Nothing
    
    
    Screen.MousePointer = vbDefault

Error1:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub



Private Function ExisteFacturaEnHco() As Boolean
'Comprobamos si la factura ya existe en la tabla de Facturas a Proveedor: scafpc
Dim Cad As String

    ExisteFacturaEnHco = False
    'Tiene que tener valor los 3 campos de clave primaria antes de comprobar
    If Not (Text1(0).Text <> "" And Text1(1).Text <> "" And Text1(3).Text <> "") Then Exit Function
    
    ' No debe existir el número de factura para el proveedor en hco
    Cad = "SELECT count(*) FROM scafpc "
    Cad = Cad & " WHERE codprove=" & Text1(3).Text & " AND numfactu=" & DBSet(Text1(0).Text, "T") & " AND year(fecfactu)=" & Year(Text1(1).Text)
    If RegistrosAListar(Cad) > 0 Then
        MsgBox "Factura de proveedor ya existente. Reintroduzca.", vbExclamation
        ExisteFacturaEnHco = True
        Exit Function
    End If
End Function

Private Sub RefrescarAlbaranes()
Dim I As Integer
Dim SQL As String
Dim Itm As ListItem
Dim Rs As ADODB.Recordset
    

    For I = 1 To ListView1.ListItems.Count
        SQL = "SELECT scaalp.numalbar,scaalp.fechaalb,scaalp.codforpa,sforpa.nomforpa,scaalp.dtoppago,scaalp.dtognral, "
        SQL = SQL & " sum(slialp.importel) as bruto "
        SQL = SQL & " FROM (scaalp LEFT OUTER JOIN sforpa ON scaalp.codforpa=sforpa.codforpa) "
        SQL = SQL & " INNER JOIN slialp ON scaalp.numalbar = slialp.numalbar  AND scaalp.fechaalb=slialp.fechaalb AND scaalp.codprove=slialp.codprove "
        SQL = SQL & " WHERE scaalp.codprove =" & Text1(3).Text & " AND scaalp.numalbar=" & DBSet(ListView1.ListItems(I).Text, "T") & " AND scaalp.fechaalb=" & DBSet(ListView1.ListItems(I).SubItems(1), "F")
        SQL = SQL & " GROUP BY scaalp.numalbar, scaalp.fechaalb, scaalp.codforpa, scaalp.dtoppago,scaalp.dtognral "
        SQL = SQL & " ORDER BY scaalp.numalbar"

        Set Rs = New ADODB.Recordset
        Rs.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText

        If Not Rs.EOF Then 'Actualizamos los datos de este item en el list
            ListView1.ListItems(I).SubItems(2) = Rs!Codforpa
            ListView1.ListItems(I).SubItems(3) = Rs!nomforpa
            ListView1.ListItems(I).SubItems(4) = Rs!DtoPPago
            ListView1.ListItems(I).SubItems(5) = Rs!DtoGnral
            ListView1.ListItems(I).SubItems(6) = Rs!Bruto

        End If
        
        If ListView1.ListItems(I).Checked Then 'comprobamos otra vez el chek y recalculamos factura
            Set Itm = ListView1.ListItems(I)
            ListView1_ItemCheck Itm
        End If

        Rs.Close
        Set Rs = Nothing
    Next I
    
    'recalcular el total de la factura
     For I = 1 To ListView1.ListItems.Count
        If ListView1.ListItems(I).Checked Then
            CalcularDatosFactura
            Exit For
        End If
     Next I
End Sub

Private Sub FechaVtoVisible(visible As Boolean)
    Text1(23).visible = visible
    Text1(23).Enabled = visible
    Label1(12).visible = visible
    Label1(12).Enabled = visible
    imgFecha(2).visible = visible
    imgFecha(2).Enabled = visible
End Sub






