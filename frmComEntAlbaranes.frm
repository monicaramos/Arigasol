VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmComEntAlbaranes 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Albaranes Proveedor"
   ClientHeight    =   6945
   ClientLeft      =   45
   ClientTop       =   4335
   ClientWidth     =   11745
   Icon            =   "frmComEntAlbaranes.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6945
   ScaleWidth      =   11745
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Height          =   840
      Left            =   135
      TabIndex        =   51
      Top             =   495
      Width           =   11415
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   1
         Left            =   1440
         MaxLength       =   10
         TabIndex        =   1
         Tag             =   "Fecha Albaran|F|N|||scaalp|fechaalb|dd/mm/yyyy|S|"
         Top             =   360
         Width           =   1185
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   25
         Left            =   3720
         MaxLength       =   10
         TabIndex        =   2
         Tag             =   "Fecha|F|N|||scaalp|fecturno|dd/mm/yyyy||"
         Top             =   360
         Width           =   1185
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   26
         Left            =   5010
         MaxLength       =   30
         TabIndex        =   3
         Tag             =   "Turno|N|N|0|9|scaalp|codturno|0||"
         Text            =   "Text1"
         Top             =   360
         Width           =   510
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   4
         Left            =   7095
         MaxLength       =   30
         TabIndex        =   4
         Tag             =   "Cod. Proveedor|N|N|0|999999|scaalp|codprove|000000|S|"
         Text            =   "Text1"
         Top             =   360
         Width           =   780
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   5
         Left            =   7920
         MaxLength       =   40
         TabIndex        =   5
         Tag             =   "Nombre Proveedor|T|N|||scaalp|nomprove||N|"
         Text            =   "Text1 wwwwwwwwwwwwwwwwwwwwwwwwwwwwwww aq"
         Top             =   360
         Width           =   3400
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000013&
         Height          =   315
         Index           =   0
         Left            =   240
         MaxLength       =   10
         TabIndex        =   0
         Tag             =   "Nº Albaran|T|N|0||scaalp|numalbar||S|"
         Text            =   "Text1 7"
         Top             =   360
         Width           =   1125
      End
      Begin VB.Label Label1 
         Caption         =   "Turno"
         Height          =   255
         Index           =   5
         Left            =   5010
         TabIndex        =   115
         Top             =   180
         Width           =   735
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   1
         Left            =   4635
         Picture         =   "frmComEntAlbaranes.frx":000C
         ToolTipText     =   "Buscar fecha"
         Top             =   150
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha"
         Height          =   255
         Index           =   4
         Left            =   3720
         TabIndex        =   114
         Top             =   180
         Width           =   855
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   0
         Left            =   6810
         Picture         =   "frmComEntAlbaranes.frx":0097
         ToolTipText     =   "Buscar proveedor"
         Top             =   360
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Proveedor"
         Height          =   255
         Index           =   0
         Left            =   5985
         TabIndex        =   54
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha Alb."
         Height          =   255
         Index           =   14
         Left            =   1440
         TabIndex        =   53
         Top             =   165
         Width           =   855
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   0
         Left            =   2355
         Picture         =   "frmComEntAlbaranes.frx":0199
         ToolTipText     =   "Buscar fecha"
         Top             =   135
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Nº Albaran"
         Height          =   255
         Index           =   50
         Left            =   240
         TabIndex        =   52
         Top             =   165
         Width           =   855
      End
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H80000018&
      Height          =   315
      Index           =   23
      Left            =   4065
      Locked          =   -1  'True
      MaxLength       =   40
      TabIndex        =   103
      Text            =   "Text2"
      Top             =   765
      Width           =   3525
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   315
      Index           =   23
      Left            =   3405
      MaxLength       =   30
      TabIndex        =   102
      Text            =   "Text1"
      Top             =   765
      Width           =   660
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H80000018&
      Height          =   315
      Index           =   17
      Left            =   9585
      Locked          =   -1  'True
      MaxLength       =   15
      TabIndex        =   100
      Text            =   "ABCDKFJADKSFJAK"
      Top             =   855
      Visible         =   0   'False
      Width           =   1725
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   21
      Left            =   5085
      MaxLength       =   30
      TabIndex        =   99
      Text            =   "Text1"
      Top             =   855
      Width           =   660
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H80000018&
      Height          =   315
      Index           =   21
      Left            =   5865
      Locked          =   -1  'True
      MaxLength       =   40
      TabIndex        =   98
      Text            =   "Text2"
      Top             =   855
      Width           =   3405
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H80000018&
      Height          =   315
      Index           =   2
      Left            =   1515
      Locked          =   -1  'True
      MaxLength       =   40
      TabIndex        =   97
      Text            =   "Text2"
      Top             =   810
      Width           =   3400
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   2
      Left            =   690
      MaxLength       =   30
      TabIndex        =   96
      Text            =   "Text1"
      Top             =   810
      Width           =   780
   End
   Begin VB.Frame Frame1 
      Height          =   475
      Index           =   0
      Left            =   135
      TabIndex        =   23
      Top             =   6270
      Width           =   2175
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
         Left            =   240
         TabIndex        =   24
         Top             =   180
         Width           =   1755
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   10485
      TabIndex        =   21
      Top             =   6375
      Width           =   1035
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   9315
      TabIndex        =   20
      Top             =   6375
      Width           =   1035
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   330
      Left            =   240
      Top             =   3960
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
      Height          =   360
      Left            =   0
      TabIndex        =   25
      Top             =   0
      Width           =   11745
      _ExtentX        =   20717
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
            Object.ToolTipText     =   "Buscar"
            Object.Width           =   1e-4
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Ver todos"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Nuevo"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Modificar"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Eliminar"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Modificacion proveedor"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Lineas Albaran"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Nº Series"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
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
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0FF&
         Height          =   285
         Index           =   50
         Left            =   9960
         MaxLength       =   15
         TabIndex        =   88
         Text            =   "Text1 7"
         Top             =   80
         Width           =   1530
      End
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
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
         Height          =   285
         Index           =   0
         Left            =   9000
         MaxLength       =   15
         TabIndex        =   87
         Text            =   "TOTAL"
         Top             =   100
         Width           =   885
      End
      Begin VB.CheckBox chkVistaPrevia 
         Caption         =   "Vista previa"
         Height          =   195
         Left            =   5880
         TabIndex        =   26
         Top             =   120
         Width           =   1215
      End
   End
   Begin MSAdodcLib.Adodc Data2 
      Height          =   330
      Left            =   240
      Top             =   4440
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
   Begin TabDlg.SSTab SSTab1 
      Height          =   4830
      Left            =   120
      TabIndex        =   27
      Tag             =   "Fecha Oferta|F|N|||scapre|fecentre|dd/mm/yyyy|N|"
      Top             =   1350
      Width           =   11415
      _ExtentX        =   20135
      _ExtentY        =   8520
      _Version        =   393216
      Style           =   1
      TabsPerRow      =   4
      TabHeight       =   520
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
      TabPicture(0)   =   "frmComEntAlbaranes.frx":0224
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1(3)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label1(35)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label1(1)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Text2(1)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "ToolAux(0)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "DataGrid1"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "txtAux(0)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "txtAux(1)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "txtAux(3)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "txtAux(4)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "txtAux(5)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "txtAux(6)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "txtAux(7)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "txtAux(2)"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "cmdAux(0)"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "cmdAux(1)"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "FrameCliente"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Text2(16)"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "Text2(0)"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).ControlCount=   19
      TabCaption(1)   =   "Otros Datos"
      TabPicture(1)   =   "frmComEntAlbaranes.frx":0240
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "FrameHco"
      Tab(1).Control(1)=   "Text1(19)"
      Tab(1).Control(2)=   "Text1(18)"
      Tab(1).Control(3)=   "Text1(17)"
      Tab(1).Control(4)=   "Text1(16)"
      Tab(1).Control(5)=   "Text1(15)"
      Tab(1).Control(6)=   "Label1(45)"
      Tab(1).ControlCount=   7
      TabCaption(2)   =   "Totales"
      TabPicture(2)   =   "frmComEntAlbaranes.frx":025C
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "FrameFactura"
      Tab(2).ControlCount=   1
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   0
         Left            =   1395
         Locked          =   -1  'True
         MaxLength       =   60
         TabIndex        =   106
         Text            =   "Text2 wwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwqa"
         Top             =   4395
         Width           =   6765
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   16
         Left            =   1395
         Locked          =   -1  'True
         MaxLength       =   60
         TabIndex        =   37
         Text            =   "Text2 wwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwqa"
         Top             =   4050
         Width           =   6765
      End
      Begin VB.Frame FrameHco 
         Caption         =   "Datos  Eliminación"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00972E0B&
         Height          =   1350
         Left            =   -74640
         TabIndex        =   89
         Top             =   405
         Width           =   5775
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   22
            Left            =   1440
            MaxLength       =   10
            TabIndex        =   92
            Top             =   345
            Width           =   1185
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   24
            Left            =   1455
            MaxLength       =   30
            TabIndex        =   91
            Text            =   "Text1"
            Top             =   780
            Width           =   660
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Height          =   315
            Index           =   24
            Left            =   2115
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   90
            Text            =   "Text2"
            Top             =   780
            Width           =   3525
         End
         Begin VB.Label Label1 
            Caption         =   "Fecha"
            Height          =   255
            Index           =   37
            Left            =   180
            TabIndex        =   94
            Top             =   345
            Width           =   1215
         End
         Begin VB.Label Label1 
            Caption         =   "Incidencia"
            Height          =   255
            Index           =   40
            Left            =   180
            TabIndex        =   93
            Top             =   810
            Width           =   825
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   7
            Left            =   1080
            Picture         =   "frmComEntAlbaranes.frx":0278
            ToolTipText     =   "Buscar incidencia"
            Top             =   780
            Width           =   240
         End
      End
      Begin VB.Frame FrameFactura 
         Height          =   3300
         Left            =   -74640
         TabIndex        =   55
         Top             =   600
         Width           =   10575
         Begin VB.TextBox Text3 
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
            Index           =   49
            Left            =   7200
            MaxLength       =   15
            TabIndex        =   72
            Text            =   "Text1 7"
            Top             =   2640
            Width           =   1845
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Height          =   285
            Index           =   48
            Left            =   7560
            MaxLength       =   15
            TabIndex        =   71
            Text            =   "Text1 7"
            Top             =   2085
            Width           =   1485
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   42
            Left            =   5040
            MaxLength       =   5
            TabIndex        =   70
            Text            =   "Text1 7"
            Top             =   2085
            Width           =   525
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   39
            Left            =   4320
            MaxLength       =   4
            TabIndex        =   69
            Text            =   "Text1 7"
            Top             =   2100
            Width           =   525
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   45
            Left            =   5760
            MaxLength       =   15
            TabIndex        =   68
            Text            =   "Text1 7"
            Top             =   2085
            Width           =   1485
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Height          =   285
            Index           =   47
            Left            =   7560
            MaxLength       =   15
            TabIndex        =   67
            Text            =   "Text1 7"
            Top             =   1755
            Width           =   1485
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   41
            Left            =   5040
            MaxLength       =   5
            TabIndex        =   66
            Text            =   "Text1 7"
            Top             =   1755
            Width           =   525
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   38
            Left            =   4320
            MaxLength       =   4
            TabIndex        =   65
            Text            =   "Text1 7"
            Top             =   1755
            Width           =   525
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   44
            Left            =   5760
            MaxLength       =   15
            TabIndex        =   64
            Text            =   "Text1 7"
            Top             =   1755
            Width           =   1485
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Height          =   285
            Index           =   46
            Left            =   7560
            MaxLength       =   15
            TabIndex        =   63
            Text            =   "Text1 7"
            Top             =   1440
            Width           =   1485
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   40
            Left            =   5040
            MaxLength       =   5
            TabIndex        =   62
            Text            =   "Text1 7"
            Top             =   1440
            Width           =   525
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   37
            Left            =   4320
            MaxLength       =   4
            TabIndex        =   61
            Text            =   "Text1 7"
            Top             =   1440
            Width           =   525
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   43
            Left            =   5760
            MaxLength       =   15
            TabIndex        =   60
            Text            =   "Text1 7"
            Top             =   1440
            Width           =   1485
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Height          =   285
            Index           =   36
            Left            =   5760
            MaxLength       =   15
            TabIndex        =   59
            Text            =   "Text1 7"
            Top             =   555
            Width           =   1485
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   35
            Left            =   3960
            MaxLength       =   15
            TabIndex        =   58
            Text            =   "Text1 7"
            Top             =   555
            Width           =   1365
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   34
            Left            =   2160
            MaxLength       =   15
            TabIndex        =   57
            Text            =   "Text1 7"
            Top             =   555
            Width           =   1365
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   33
            Left            =   240
            MaxLength       =   15
            TabIndex        =   56
            Text            =   "Text1 7"
            Top             =   555
            Width           =   1485
         End
         Begin VB.Label Label1 
            Caption         =   "Cod. IVA"
            Height          =   255
            Index           =   42
            Left            =   4320
            TabIndex        =   86
            Top             =   1200
            Width           =   735
         End
         Begin VB.Label Label1 
            Caption         =   "% IVA"
            Height          =   255
            Index           =   41
            Left            =   5040
            TabIndex        =   85
            Top             =   1230
            Width           =   495
         End
         Begin VB.Label Label1 
            Caption         =   "TOTAL ALBARAN"
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
            Height          =   285
            Index           =   39
            Left            =   5640
            TabIndex        =   84
            Top             =   2660
            Width           =   1530
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
            TabIndex        =   83
            Top             =   2160
            Width           =   135
         End
         Begin VB.Line Line1 
            X1              =   4320
            X2              =   7320
            Y1              =   1065
            Y2              =   1065
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
            Index           =   8
            Left            =   7320
            TabIndex        =   82
            Top             =   1320
            Width           =   135
         End
         Begin VB.Label Label1 
            Caption         =   "Imp. IVA"
            Height          =   255
            Index           =   33
            Left            =   7560
            TabIndex        =   81
            Top             =   1230
            Width           =   1335
         End
         Begin VB.Label Label1 
            Caption         =   "="
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
            Index           =   32
            Left            =   5520
            TabIndex        =   80
            Top             =   480
            Width           =   135
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
            Index           =   31
            Left            =   3720
            TabIndex        =   79
            Top             =   480
            Width           =   135
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
            Index           =   30
            Left            =   1920
            TabIndex        =   78
            Top             =   480
            Width           =   135
         End
         Begin VB.Label Label1 
            Caption         =   "Base Imponible"
            Height          =   255
            Index           =   18
            Left            =   5760
            TabIndex        =   77
            Top             =   360
            Width           =   1215
         End
         Begin VB.Label Label1 
            Caption         =   "Imp. Dto Gn"
            Height          =   255
            Index           =   22
            Left            =   3960
            TabIndex        =   76
            Top             =   360
            Width           =   855
         End
         Begin VB.Label Label1 
            Caption         =   "Imp. Dto PP"
            Height          =   255
            Index           =   23
            Left            =   2160
            TabIndex        =   75
            Top             =   360
            Width           =   855
         End
         Begin VB.Label Label1 
            Caption         =   "Imp. Bruto"
            Height          =   255
            Index           =   24
            Left            =   240
            TabIndex        =   74
            Top             =   360
            Width           =   855
         End
         Begin VB.Label Label1 
            Caption         =   "Base Imponible"
            Height          =   255
            Index           =   27
            Left            =   5760
            TabIndex        =   73
            Top             =   1200
            Width           =   1215
         End
      End
      Begin VB.Frame FrameCliente 
         Height          =   1245
         Left            =   225
         TabIndex        =   41
         Top             =   315
         Width           =   10935
         Begin VB.TextBox Text1 
            Height          =   315
            Index           =   11
            Left            =   6885
            MaxLength       =   30
            TabIndex        =   11
            Tag             =   "Provincia|T|N|||scaalp|proprove||N|"
            Text            =   "Text1 Text1 Text1 Text1 Text22"
            Top             =   855
            Width           =   2445
         End
         Begin VB.TextBox Text1 
            Height          =   315
            Index           =   9
            Left            =   1245
            MaxLength       =   6
            TabIndex        =   9
            Tag             =   "CPostal|T|N|||scaalp|codpobla||N|"
            Text            =   "Text15"
            Top             =   870
            Width           =   630
         End
         Begin VB.TextBox Text1 
            Height          =   315
            Index           =   10
            Left            =   1875
            MaxLength       =   30
            TabIndex        =   10
            Tag             =   "Población|T|N|||scaalp|pobprove||N|"
            Text            =   "Text1 wwwwwwwwwwwwwwwwwwwwwwaq"
            Top             =   870
            Width           =   3405
         End
         Begin VB.TextBox Text1 
            Height          =   315
            Index           =   7
            Left            =   3435
            MaxLength       =   20
            TabIndex        =   7
            Tag             =   "teléfono Proveedor|T|S|||scaalp|telprove||N|"
            Text            =   "12345678911234567899"
            Top             =   150
            Width           =   1845
         End
         Begin VB.TextBox Text1 
            Height          =   315
            Index           =   6
            Left            =   1245
            MaxLength       =   15
            TabIndex        =   6
            Tag             =   "NIF Proveedor|T|N|||scaalp|nifprove||N|"
            Text            =   "123456789"
            Top             =   150
            Width           =   1350
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   12
            Left            =   6885
            MaxLength       =   30
            TabIndex        =   12
            Tag             =   "Forma de Pago|N|N|0|999|scaalp|codforpa|000|N|"
            Text            =   "Text1"
            Top             =   150
            Width           =   540
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Height          =   315
            Index           =   12
            Left            =   7470
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   43
            Text            =   "Text2"
            Top             =   150
            Width           =   3390
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   13
            Left            =   6885
            MaxLength       =   7
            TabIndex        =   13
            Tag             =   "Descuento P.Pago|N|N|0|99.90|scaalp|dtoppago|#0.00|N|"
            Text            =   "Text1 7"
            Top             =   510
            Width           =   525
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   14
            Left            =   8445
            MaxLength       =   7
            TabIndex        =   14
            Tag             =   "Descuento General|N|N|0|99.90|scaalp|dtognral|#0.00|N|"
            Text            =   "Text1 7"
            Top             =   510
            Width           =   525
         End
         Begin VB.TextBox Text1 
            Height          =   315
            Index           =   8
            Left            =   1245
            MaxLength       =   35
            TabIndex        =   8
            Tag             =   "Domicilio|T|N|||scaalp|domprove||N|"
            Text            =   "Text1 wwwwwwwwwwwwwwwwwwwwwwwwww aq"
            Top             =   510
            Width           =   4030
         End
         Begin VB.Image imgBuscar 
            Enabled         =   0   'False
            Height          =   240
            Index           =   5
            Left            =   960
            Picture         =   "frmComEntAlbaranes.frx":037A
            ToolTipText     =   "Buscar proveedor vario"
            Top             =   195
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Provincia"
            Height          =   255
            Index           =   17
            Left            =   5715
            TabIndex        =   50
            Top             =   915
            Width           =   735
         End
         Begin VB.Label Label1 
            Caption         =   "Población"
            Height          =   255
            Index           =   16
            Left            =   120
            TabIndex        =   49
            Top             =   870
            Width           =   735
         End
         Begin VB.Label Label1 
            Caption         =   "Teléfono"
            Height          =   255
            Index           =   19
            Left            =   2745
            TabIndex        =   48
            Top             =   150
            Width           =   735
         End
         Begin VB.Label Label1 
            Caption         =   "NIF"
            Height          =   255
            Index           =   20
            Left            =   120
            TabIndex        =   47
            Top             =   150
            Width           =   615
         End
         Begin VB.Label Label1 
            Caption         =   "Forma Pago"
            Height          =   255
            Index           =   15
            Left            =   5700
            TabIndex        =   46
            Top             =   150
            Width           =   855
         End
         Begin VB.Label Label1 
            Caption         =   "Dto. P. Pago"
            Height          =   255
            Index           =   25
            Left            =   5700
            TabIndex        =   45
            Top             =   510
            Width           =   975
         End
         Begin VB.Label Label1 
            Caption         =   "Dto. Gral"
            Height          =   255
            Index           =   26
            Left            =   7740
            TabIndex        =   44
            Top             =   510
            Width           =   735
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   3
            Left            =   6600
            Picture         =   "frmComEntAlbaranes.frx":047C
            ToolTipText     =   "Buscar forma de pago"
            Top             =   150
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Domicilio"
            Height          =   255
            Index           =   7
            Left            =   120
            TabIndex        =   42
            Top             =   510
            Width           =   735
         End
      End
      Begin VB.CommandButton cmdAux 
         Appearance      =   0  'Flat
         Caption         =   "+"
         Height          =   315
         Index           =   1
         Left            =   2640
         TabIndex        =   40
         ToolTipText     =   "Buscar artículo"
         Top             =   3540
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.CommandButton cmdAux 
         Appearance      =   0  'Flat
         Caption         =   "+"
         Height          =   315
         Index           =   0
         Left            =   960
         TabIndex        =   39
         ToolTipText     =   "Buscar almacen"
         Top             =   3540
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.TextBox txtAux 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   2
         Left            =   2880
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   31
         Tag             =   "Nombre Artículo"
         Text            =   "nomArtic"
         Top             =   3600
         Visible         =   0   'False
         Width           =   3285
      End
      Begin VB.TextBox txtAux 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000013&
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   7
         Left            =   9720
         MaxLength       =   12
         TabIndex        =   36
         Tag             =   "Importe"
         Text            =   "Importe"
         Top             =   3600
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox txtAux 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   6
         Left            =   9120
         MaxLength       =   30
         TabIndex        =   35
         Tag             =   "Descuento 2"
         Text            =   "Dto2"
         Top             =   3600
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox txtAux 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000013&
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   5
         Left            =   8520
         MaxLength       =   5
         TabIndex        =   34
         Tag             =   "Descuento 1"
         Text            =   "Dto1"
         Top             =   3600
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox txtAux 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   4
         Left            =   7440
         MaxLength       =   12
         TabIndex        =   33
         Tag             =   "Precio"
         Text            =   "123,456.7879"
         Top             =   3600
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox txtAux 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000013&
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   3
         Left            =   6240
         MaxLength       =   16
         TabIndex        =   32
         Tag             =   "Cantidad"
         Text            =   "1,234,567,891.25"
         Top             =   3600
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox txtAux 
         Appearance      =   0  'Flat
         BackColor       =   &H80000013&
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   1
         Left            =   1200
         MaxLength       =   18
         TabIndex        =   30
         Tag             =   "Código Artículo"
         Text            =   "Artic Artic Artic5"
         Top             =   3540
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.TextBox txtAux 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   0
         Left            =   360
         MaxLength       =   15
         TabIndex        =   29
         Tag             =   "Código Almacen"
         Text            =   "codalmac"
         Top             =   3540
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox Text1 
         Height          =   300
         Index           =   19
         Left            =   -74640
         MaxLength       =   80
         TabIndex        =   19
         Tag             =   "Observación 5|T|S|||scaalp|observa5||N|"
         Top             =   3360
         Width           =   8445
      End
      Begin VB.TextBox Text1 
         Height          =   300
         Index           =   18
         Left            =   -74640
         MaxLength       =   80
         TabIndex        =   18
         Tag             =   "Observación 4|T|S|||scaalp|observa4||N|"
         Top             =   3060
         Width           =   8445
      End
      Begin VB.TextBox Text1 
         Height          =   300
         Index           =   17
         Left            =   -74640
         MaxLength       =   80
         TabIndex        =   17
         Tag             =   "Observación 3|T|S|||scaalp|observa3||N|"
         Top             =   2760
         Width           =   8445
      End
      Begin VB.TextBox Text1 
         Height          =   300
         Index           =   16
         Left            =   -74640
         MaxLength       =   80
         TabIndex        =   16
         Tag             =   "Observación 2|T|S|||scaalp|observa2||N|"
         Top             =   2460
         Width           =   8445
      End
      Begin VB.TextBox Text1 
         Height          =   300
         Index           =   15
         Left            =   -74640
         MaxLength       =   80
         TabIndex        =   15
         Tag             =   "Observación 1|T|S|||scaalp|observa1||N|"
         Top             =   2160
         Width           =   8445
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "frmComEntAlbaranes.frx":057E
         Height          =   2025
         Left            =   240
         TabIndex        =   38
         Top             =   1980
         Width           =   10935
         _ExtentX        =   19288
         _ExtentY        =   3572
         _Version        =   393216
         AllowUpdate     =   -1  'True
         AllowArrows     =   -1  'True
         ColumnHeaders   =   -1  'True
         HeadLines       =   1
         RowHeight       =   15
         FormatLocked    =   -1  'True
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
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.Toolbar ToolAux 
         Height          =   390
         Index           =   0
         Left            =   225
         TabIndex        =   105
         Top             =   1575
         Width           =   1110
         _ExtentX        =   1958
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
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   1
         Left            =   9030
         Locked          =   -1  'True
         MaxLength       =   60
         TabIndex        =   108
         Text            =   "Text2 wwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwqa"
         Top             =   3465
         Width           =   2175
      End
      Begin VB.Label Label1 
         Caption         =   "Familia"
         Height          =   255
         Index           =   1
         Left            =   495
         TabIndex        =   107
         Top             =   4455
         Width           =   570
      End
      Begin VB.Label Label1 
         Caption         =   "Ampliación"
         Height          =   255
         Index           =   35
         Left            =   495
         TabIndex        =   95
         Top             =   4095
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Observaciones"
         Height          =   255
         Index           =   45
         Left            =   -74640
         TabIndex        =   28
         Top             =   1920
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Unidad"
         Height          =   255
         Index           =   3
         Left            =   8400
         TabIndex        =   109
         Top             =   3525
         Width           =   615
      End
   End
   Begin VB.CommandButton cmdRegresar 
      Caption         =   "&Regresar"
      Height          =   375
      Left            =   10485
      TabIndex        =   22
      Top             =   6345
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   3
      Left            =   2880
      MaxLength       =   7
      TabIndex        =   110
      Tag             =   "Nº Pedido|N|S|0||scaalp|numpedpr|0000000|N|"
      Text            =   "Text1 7"
      Top             =   885
      Visible         =   0   'False
      Width           =   885
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   20
      Left            =   3855
      MaxLength       =   10
      TabIndex        =   111
      Tag             =   "Fecha Pedido|F|S|||scaalp|fecpedpr|dd/mm/yyyy|N|"
      Top             =   885
      Width           =   1185
   End
   Begin VB.Image imgBuscar 
      Enabled         =   0   'False
      Height          =   240
      Index           =   2
      Left            =   5130
      Picture         =   "frmComEntAlbaranes.frx":0593
      ToolTipText     =   "Buscar población"
      Top             =   5445
      Width           =   240
   End
   Begin VB.Image imgBuscar 
      Height          =   240
      Index           =   6
      Left            =   3030
      Picture         =   "frmComEntAlbaranes.frx":0695
      ToolTipText     =   "Buscar trabajador"
      Top             =   765
      Width           =   240
   End
   Begin VB.Label Label1 
      Caption         =   "Trabajador"
      Height          =   255
      Index           =   38
      Left            =   2070
      TabIndex        =   104
      Top             =   765
      Width           =   825
   End
   Begin VB.Label Label1 
      Caption         =   "Nº Lote"
      Height          =   210
      Index           =   2
      Left            =   9405
      TabIndex        =   101
      Top             =   900
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Image imgBuscar 
      Height          =   240
      Index           =   4
      Left            =   6165
      Picture         =   "frmComEntAlbaranes.frx":0797
      ToolTipText     =   "Buscar trabajador"
      Top             =   630
      Width           =   240
   End
   Begin VB.Image imgBuscar 
      Height          =   240
      Index           =   1
      Left            =   405
      Picture         =   "frmComEntAlbaranes.frx":0899
      ToolTipText     =   "Buscar trabajador"
      Top             =   825
      Width           =   240
   End
   Begin VB.Label Label1 
      Caption         =   "Fecha Pedido"
      Height          =   255
      Index           =   10
      Left            =   3855
      TabIndex        =   113
      Top             =   690
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Nº Pedido"
      Height          =   255
      Index           =   11
      Left            =   2880
      TabIndex        =   112
      Top             =   690
      Visible         =   0   'False
      Width           =   735
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
      Begin VB.Menu mnLineas 
         Caption         =   "&Lineas"
         Enabled         =   0   'False
         HelpContextID   =   2
         Shortcut        =   ^L
         Visible         =   0   'False
      End
      Begin VB.Menu mnSalir 
         Caption         =   "&Salir"
         Shortcut        =   ^S
      End
   End
End
Attribute VB_Name = "frmComEntAlbaranes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public DatosADevolverBusqueda As String    'Tendra el nº de text que quiere que devuelva, empipados
Public Event DatoSeleccionado(CadenaSeleccion As String)

'Si se llama de la busqueda en el frmAlmMovimArticulos se accede
'a las tablas del Albaran o de Facturas de movimiento seleccionado (solo consulta)
Public hcoCodMovim As String 'cod. movim
Public hcoFechaMovim As Date 'Fecha del Movim
Public hcoCodProve As Integer 'Codigo de Proveedor

Public EsHistorico As Boolean 'Si es true abrir el formulario con la tabla de
                              'de historico schalb, y solo en modo de consulta
                              
'cadena que selecciona los albaranes de un proveedor para mostrar
'antes de facturarlos
Public cadSelAlbaranes As String

Private WithEvents frmB As frmBuscaGrid 'Form para busquedas
Attribute frmB.VB_VarHelpID = -1
Private WithEvents frmF As frmCal 'Form Calendario Fecha
Attribute frmF.VB_VarHelpID = -1

Private WithEvents frmProve As frmManProve  'Form Mto Proveedores
Attribute frmProve.VB_VarHelpID = -1
'Private WithEvents frmPV As frmComProveV   'Form Mto Proveedores Varios

Private WithEvents frmFP As frmManFpago 'Form Mto Formas de Pago
Attribute frmFP.VB_VarHelpID = -1
'Private WithEvents frmAlm As frmManAlmProp   'Form Almacenes Propios
Private WithEvents frmArt As frmManArtic   'Form Articulos
Attribute frmArt.VB_VarHelpID = -1

'Private WithEvents frmMen As frmMensajes  'Form Mensajes
Private WithEvents frmList As frmListadoOfer
Attribute frmList.VB_VarHelpID = -1

'-------------------------------------------------------------------------
Private Modo As Byte
'-----------------------------
'Se distinguen varios Modos
'   0.-  Formulario limpio sin nungun campo rellenado
'   1.-  Preparando para hacer la busqueda
'   2.-  Ya tenemos registros y los vamos a recorrer
'        y podemos editarlos Edicion del campo
'   3.-  Insercion de nuevo registro
'   4.-  Modificar
'   5.-  Mantenimiento Lineas
'-------------------------------------------------------------------------

Dim ModificaLineas As Byte
'1.- Añadir,  2.- Modificar,  3.- Borrar,  0.-Pasar control a Lineas
'4.- Mantenimiento de Nº de Serie

Dim TituloLinea As String 'Descripcion de la linea que estamos en Mantenimiento

Dim PrimeraVez As Boolean


Dim CodTipoMov As String
'Codigo tipo de movimiento en función del valor en la tabla de parámetros: stipom

Dim EsDeVarios As Boolean
'Si el Proveedor mostrado es de Varios o No

'SQL de la tabla principal del formulario
Private CadenaConsulta As String

Private Ordenacion As String
Private NombreTabla As String  'Nombre de la tabla o de la
Private NomTablaLineas As String 'Nombre de la Tabla de lineas
Private kCampo As Integer
'-------------------------------------------------------------------------
Private HaDevueltoDatos As Boolean

Dim btnAnyadir As Byte
'Variable que indica el número del Boton  Anyadir en la Toolbar1
Dim btnPrimero As Byte
'Variable que indica el número del Boton  PrimerRegistro en la Toolbar1

Dim indice As Integer


Private VieneDeBuscar As Boolean
'Para cuando devuelve dos poblaciones con el mismo codigo Postal. Si viene de pulsar prismatico
'de busqueda poner el valor de poblacion seleccionado y no volver a recuperar de la Base de Datos
Private HaCambiadoCP As Boolean
'Para saber si tras haber vuelto de prismaticos ha cambiado el valor del CPostal

Dim cadList As String 'cadena para pasar al historico



Private Sub cmdAceptar_Click()
Dim PrimeraLin As Boolean 'Si se inserta la primera linea no esta creado el datagrid1 entonces llamar
                          ' a DataGrid, sino llamar solo a DataGrid2
Dim numlinea As String
Dim PrecioUC As Currency

    Screen.MousePointer = vbHourglass
    On Error GoTo Error1

    Select Case Modo
        Case 1  'BUSQUEDA
            HacerBusqueda
            
        Case 3 'INSERTAR CABECERA
            If DatosOk Then InsertarCabecera
            
        Case 4  'MODIFICAR CABECERA
            If DatosOk Then
                If ModificarCabAlbaran Then
                    If cadSelAlbaranes = "" Then TerminaBloquear
                    PosicionarData
                End If
            End If
            
         Case 5 'InsertarModificar LINEAS
            'Antes de insertar la linea guardamos el sartic.preciouc actual
            'para aplicar margen despues, pq en Insertar linea se actualiza ya el preciouc
            PrecioUC = ComprobarCero(DevuelveDesdeBDNew(cPTours, "sartic", "ultpreci", "codartic", txtAux(1).Text, "T"))
                
         
            'Actualizar el registro en la tabla de lineas 'slialb'
            If ModificaLineas = 1 Then 'INSERTAR lineas Albaran
                PrimeraLin = False
                If Data2.Recordset.EOF = True Then PrimeraLin = True
                          
                If InsertarLinea(numlinea) Then
                    'Comprobar si Hay Nº SERIE en compras y Mostrar
                    'ventana para pedir los Nº Serie de la cantidad introducida
'--monica
'                    If vParamAplic.NumSeries Then
'                        ComprobarNumSeries (numlinea)
'                    End If

'--monica: no tenemos slista
'                    'Comprobar si se ha modificado el precio desde la ultima compra
'                    'y preguntar quiere modificar el PVP del articulo aplicandole su margen
'                    'y el precio de las TArifas aplicandole el margen
'                    '-- Laura 19/12/2006: el precio de compra es el precio con los descuentos (importe/cantidad)
'                    'If precioUC <> CCur(txtAux(4).Text) Then
'                    If precioUC <> Round2(CCur(txtAux(7).Text) / CCur(txtAux(3).Text), 4) Then
'                        If MsgBox("Se ha modificado el precio última compra." & vbCrLf & "¿Desea actualizar los precios de venta?", vbQuestion + vbYesNo) = vbYes Then
'                            'Comprobar que el artículo tiene margen comercial
'                            If ArticuloTieneMargen(txtAux(1).Text) Then
'                                'bloquear las tarifas del articulo para modificar
''                                If BloqueaRegistro("slista", "codartic=" & DBSet(txtAux(1).Text, "T")) Then
'                                    'Aplicar margen comercial a los precios
'                                    'Modificar precios de venta en articulo y tarifas
'                                    frmComActPrecios.parCodArtic = txtAux(1).Text
'                                    frmComActPrecios.parNomArtic = txtAux(2).Text
'        '                            frmcomactprecios.parPrecioUC =
'                                    frmComActPrecios.Show vbModal
''                                End If
'                            End If
'                        End If
'                    End If
                    
                    
                    If PrimeraLin Then
                        CargaGrid DataGrid1, Data2, True
                    Else
                        CargaGrid2 DataGrid1, Data2
                    End If
                    BotonAnyadirLinea
                End If
                
            ElseIf ModificaLineas = 2 Then 'MODIFICAR lineas
                If DatosOkLinea() Then
                    If ModificarLinea Then
                        'Comprobar si Hay Nº SERIE en compras
                        'si new_cantidad>old_cantidad pedir los + nº serie
                        
                        'si new_cantidad<old_cantidad mostrar los comprados para quitar los q no queremos
'                        numlinea = Data2.Recordset!numlinea
'                        If vParamAplic.NumSeries Then
'                            ComprobarNumSeries (numlinea)
'                        End If
'
                        
'--monica: no tenemos slista
'                        'Comprobar si se ha modificado el precio desde la ultima compra
'                        'y preguntar quiere modificar el PVP del articulo aplicandole su margen
'                        'y el precio de las TArifas aplicandole el margen
'                        '-- Laura 19/12/2006: el precio de compra es el precio con los descuentos (importe/cantidad)
'                        'If precioUC <> CCur(txtAux(4).Text) Then
'                        If precioUC <> Round2(CCur(txtAux(7).Text) / CCur(txtAux(3).Text), 4) Then
'                            If MsgBox("Se ha modificado el precio última compra." & vbCrLf & "¿Desea actualizar los precios de venta?", vbQuestion + vbYesNo) = vbYes Then
'                                'Comprobar que el artículo tiene margen comercial
'                                If ArticuloTieneMargen(txtAux(1).Text) Then
'                                    'Aplicar margen comercial a los precios
'                                    'Modificar precios de venta en articulo y tarifas
'                                    frmComActPrecios.parCodArtic = txtAux(1).Text
'                                    frmComActPrecios.parNomArtic = txtAux(2).Text
'        '                            frmcomactprecios.parPrecioUC =
'                                    frmComActPrecios.Show vbModal
'                                End If
'                            End If
'                        End If

'--monica: rollo toolbar
'                        TerminaBloquear
                        CargaGrid2 DataGrid1, Data2
                        ModificaLineas = 0
                        CargaTxtAux False, False
'--monica:rollo toolbar
'                        PonerBotonCabecera True
                        BloquearTxt text2(16), True
                        BloquearTxt text2(17), True
                        
'++monica. rollo toolbar
                        PonerModo 2
                        PonerCampos
                    End If
                    Me.DataGrid1.Enabled = True
                End If
            End If
            CalcularDatosFactura 'rellenar campos pestaña de totales
    End Select
    Screen.MousePointer = vbDefault

Error1:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub


Private Function ComprobarCambioFecha() As Boolean
'Comprueba si se ha modificado el campo fecha de la cabecera.
'Ya que es clave primaria y se deberan cambiar tambien la fecha
'en tablas sliap y smoval
Dim RS As ADODB.Recordset
Dim sql As String
Dim Izquierda As String, Derecha As String
Dim llis As Collection
Dim i As Integer
Dim b As Boolean


    If Data1.Recordset.EOF Then Exit Function

    
    If (CDate(Text1(1).Text) <> CDate(Data1.Recordset!fechaalb)) Then
    'si ha modificado la fecha de albaran
        On Error GoTo EComprobar
        
        'seleccionar todas las lineas de ese albaran para actualizar la fecha (slialp)
        sql = "SELECT * FROM " & NomTablaLineas & " WHERE numalbar=" & DBSet(Data1.Recordset!numalbar, "T")
        sql = sql & " AND fechaalb=" & DBSet(Data1.Recordset!fechaalb, "F")
        sql = sql & " AND codprove=" & Data1.Recordset!CodProve
        
        Set RS = New ADODB.Recordset
        RS.Open sql, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
        
        Set llis = New Collection
            
        'Nos guardamos todas las lineas con la modificacion de la fecha para
        'volverlas a insertar
        BACKUP_TablaIzquierda RS, Izquierda
        
        While Not RS.EOF
            BACKUP_Tabla RS, Derecha, "fechaalb", CStr(Text1(1).Text)
            llis.Add Derecha
            RS.MoveNext
        Wend
        
        RS.Close
        Set RS = Nothing
        
        'Eliminamos las lineas que tenia ese albaran (slialp) para volverlas a insertar con la fecha nueva
        sql = "DELETE from slialp WHERE numalbar = " & DBSet(Data1.Recordset!numalbar, "T")
        sql = sql & " AND fechaalb=" & DBSet(Data1.Recordset!fechaalb, "F")
        sql = sql & " AND codprove=" & Data1.Recordset!CodProve
        Conn.Execute sql
        
        'Actualizamos la fecha en la cabecera (scaalp)
        sql = "UPDATE scaalp SET fechaalb = " & DBSet(Text1(1).Text, "F")
        sql = sql & " WHERE numalbar = " & DBSet(Data1.Recordset!numalbar, "T")
        sql = sql & " AND fechaalb=" & DBSet(Data1.Recordset!fechaalb, "F")
        sql = sql & " AND codprove=" & Data1.Recordset!CodProve
        Conn.Execute sql
        
        'Actualizamos la fecha en la tabla smoval
        sql = "UPDATE smoval SET fechamov=" & DBSet(Text1(1).Text, "F")
        sql = sql & " WHERE document = " & DBSet(Data1.Recordset!numalbar, "T")
        sql = sql & " AND fechamov=" & DBSet(Data1.Recordset!fechaalb, "F")
        sql = sql & " AND codigope=" & Data1.Recordset!CodProve
        sql = sql & " AND detamovi='" & CodTipoMov & "'"
        Conn.Execute sql
        
'--monica
'        'Actualizar la fecha compra en los numeros de serie del albaran (si tiene articulos con num. serie)
'        SQL = "UPDATE sserie SET fechacom=" & DBSet(Text1(1).Text, "F")
'        SQL = SQL & " WHERE fechacom=" & DBSet(Data1.Recordset!FechaAlb, "F") & " AND "
'        SQL = SQL & " numalbpr=" & DBSet(Data1.Recordset!numalbar, "T")
'        SQL = SQL & " AND codprove=" & Data1.Recordset!codProve
'        Conn.Execute SQL
        
        
        'Volvemos a insertar las lineas con la fecha correcta (slialp)
        sql = ""
        For i = 1 To llis.Count
            If (i Mod 10) = 0 Then
                sql = sql & CStr(llis(i)) & ","
                sql = Mid(sql, 1, Len(sql) - 1) 'quitamos ultima coma
                sql = "INSERT INTO " & NomTablaLineas & " " & Izquierda & " VALUES " & sql & ";"
                Conn.Execute sql
                sql = ""
            Else
                sql = sql & CStr(llis(i)) & ","
            End If
        Next i
        
        If sql <> "" Then
            sql = Mid(sql, 1, Len(sql) - 1) 'quitamos ultima coma
            sql = "INSERT INTO " & NomTablaLineas & " " & Izquierda & " VALUES " & sql & ";"
            Conn.Execute sql
            sql = ""
        End If
        Set llis = Nothing
    End If
    b = True
    
EComprobar:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "El campo fecha no se ha podido modificar", Err.Description
        b = False
    End If
    If b Then
        ComprobarCambioFecha = True
    Else
        ComprobarCambioFecha = False
    End If
End Function



Private Sub cmdAux_Click(Index As Integer)
    Select Case Index
        Case 0 'Busqueda de Cod. Almacen
'            Set frmAlm = New frmManAlmProp
'            frmAlm.DatosADevolverBusqueda = "0|"
'            frmAlm.Show vbModal
'            Set frmAlm = Nothing
        Case 1 'Busqueda de Cod. Artic
            Set frmArt = New frmManArtic
            frmArt.DatosADevolverBusqueda = "0|1|" 'Poner en modo búsqueda
            frmArt.Show vbModal
            Set frmArt = Nothing
    End Select
    PonerFoco txtAux(Index)
End Sub


Private Sub cmdCancelar_Click()
    Select Case Modo
        Case 1, 3 'Busqueda, Insertar
            LimpiarCampos
            'Poner los grid sin apuntar a nada
            LimpiarDataGrids
            PonerModo 0
            PonerFoco Text1(0)
        Case 4  'Modificar
            TerminaBloquear
            PonerModo 2
            PonerCampos
            PonerFoco Text1(0)
        Case 5 'Lineas Detalle
            TerminaBloquear
            CargaTxtAux False, False
            BloquearTxt text2(16), True
            BloquearTxt text2(17), True
            If ModificaLineas = 1 Then 'INSERTAR
                ModificaLineas = 0
                DataGrid1.AllowAddNew = False
                If Not Data2.Recordset.EOF Then Data2.Recordset.MoveFirst
            End If
            ModificaLineas = 0
'--monica: rollo toolbar
'            PonerBotonCabecera True
            Me.DataGrid1.Enabled = True
            PonerModo 2
            PonerCampos
            
    End Select
End Sub


Private Sub BotonAnyadir()
'Añadir registro en tabla de cabecera de Albaranes: scaalp (Cabecera)
Dim NomTraba As String

    LimpiarCampos 'Vacía los TextBox
    'Poner los grid sin apuntar a nada
    LimpiarDataGrids
    
    'Añadiremos el boton de aceptar y demas objetos para insertar
    PonerModo 3

    'Poner el nombre del trabajador que esta conectado
'    Text1(2).Text = PonerTrabajadorConectado(NomTraba)
'    Text2(2).Text = NomTraba
    
    Text1(1).Text = Format(Now, "dd/mm/yyyy") 'Fecha Albaran
    Text1(25).Text = Text1(1).Text
    PonerFoco Text1(0)
End Sub


Private Sub BotonAnyadirLinea()
    'Si no estaba modificando lineas salimos
    ' Es decir, si estaba insertando linea no podemos hacer otra cosa
    If ModificaLineas = 2 Then Exit Sub
       
    ModificaLineas = 1 'Ponemos Modo Añadir Linea
    'Añadiremos el boton de aceptar y demas objetos para insertar
    
'--monica: rollo toolbar
'    PonerBotonCabecera False
'    lblIndicador.Caption = "INSERTAR"
    
    AnyadirLinea DataGrid1, Data2
    CargaTxtAux True, True
    'Poner el Almacen por defecto del Trabajador
'--monica
'    txtAux(0).Text = DevuelveDesdeBDNew(conAri, "straba", "codalmac", "codtraba", Text1(2).Text, "N")
'++monica
    txtAux(0).Text = 1
    If txtAux(0).Text <> "" Then txtAux(0).Text = Format(txtAux(0).Text, "000")
    'Campo Ampliacion Linea
    text2(16).Text = ""
    text2(17).Text = ""
    text2(0).Text = ""
    text2(1).Text = ""
    BloquearTxt text2(16), False
    BloquearTxt text2(17), True
    
    PonerFoco txtAux(1)
    Me.DataGrid1.Enabled = False
End Sub


Private Sub BotonBuscar()
    'Buscar
    If Modo <> 1 Then
        LimpiarCampos
        'Poner los grid sin apuntar a nada
        LimpiarDataGrids
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


Private Sub BotonVerTodos()
    If chkVistaPrevia.Value = 1 Then
        MandaBusquedaPrevia cadSelAlbaranes
    Else
        LimpiarCampos
        LimpiarDataGrids
        If cadSelAlbaranes = "" Then
            CadenaConsulta = "Select * from " & NombreTabla & " " & Ordenacion
        Else
            CadenaConsulta = "Select * from " & NombreTabla & " " & " WHERE " & cadSelAlbaranes & Ordenacion
        End If
        PonerCadenaBusqueda
    End If
End Sub


Private Sub Desplazamiento(Index As Integer)
'Botones de Desplazamiento de la Toolbar
'Para desplazarse por los registros de control Data
    If Data1.Recordset.EOF Then Exit Sub
    DesplazamientoData Data1, Index
    PonerCampos
End Sub


Private Sub BotonModificar()
Dim sql As String
Dim DeVarios As Boolean

    On Error GoTo EModificar

    'Añadiremos el boton de aceptar y demas objetos para insertar
    PonerModo 4
    
    PonerFoco Text1(2)
    
    If EsDeVarios Then
        If Data1.Recordset.EOF Then Exit Sub
        sql = " SELECT * FROM sprvar WHERE nifprove='" & Data1.Recordset!NIFProve & "' FOR UPDATE "
        Conn.Execute sql
    End If
    
    'Si es Cliente de Varios no se pueden modificar sus datos
    DeVarios = False 'EsProveedorVarios(Text1(4).Text)
    BloquearDatosProve (DeVarios)
    
EModificar:
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub BotonModificarLinea()
'Modificar una linea
Dim vWhere As String
Dim cArt As CArticulo

    On Error GoTo EModificarLinea

    'Si no estaba modificando lineas salimos
    'Es decir, si estaba insertando linea no podemos hacer otra cosa
    If ModificaLineas = 1 Then Exit Sub '1= Insertar
    
    If Data2.Recordset.EOF Then Exit Sub
    
    vWhere = ObtenerWhereCP(False) & " and numlinea=" & Data2.Recordset!numlinea
    vWhere = Replace(vWhere, NombreTabla, NomTablaLineas)
    
    If Not BloqueaRegistro(NomTablaLineas, vWhere) Then Exit Sub
    
    ModificaLineas = 2 'Modificar
    CargaTxtAux True, False
    'ModificaLineas = 2 'Modificar
    'Añadiremos el boton de aceptar y demas objetos para insertar
'--monica: rollo toolbar
'    Me.lblIndicador.Caption = "MODIFICAR"
'    PonerBotonCabecera False
    BloquearTxt text2(16), False 'Campo Ampliacion Linea
    BloquearTxt txtAux(2), True 'campo nombre articulo
    
    'bloquear el num_lote si el articulo es de una categoria q no lleva control
    'de nº de lote
'    BloquearTxt Text2(17), (DBLet(Data2.Recordset!numlotes, "T") = "")
    Set cArt = New CArticulo
    If cArt.LeerDatos(Data2.Recordset!codartic) Then
        BloquearTxt text2(17), Not cArt.TieneNumLote
    End If
    Set cArt = Nothing
    
    
    
    PonerFoco txtAux(0)
    Me.DataGrid1.Enabled = False
    
EModificarLinea:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub


Private Sub BotonEliminar()
'Eliminar Registro de la Cabecera: Tabla de albaranes compras (scaalp)
' y los registros correspondientes de las tablas de lineas (slialp)
'al eliminar un albaran ademas habrá que restaurar valores:
' - actualizar stock en (salmac)
' - eliminar los movimientos que inserto el albaran en (smoval)
' - actualizar el ultprecio compra y ultima fecha compra en funcion del ult. movimiento ALC en smoval
' - reestablecer el precio medio ponderado
Dim cad As String

    On Error GoTo EEliminar

    'Ciertas comprobaciones
    If Data1.Recordset.EOF Then Exit Sub
    
    cad = "Cabecera de Albaranes Compras" & vbCrLf
    cad = cad & "-------------------------------------------------" & vbCrLf & vbCrLf
    cad = cad & "Va a eliminar el Albaran:            "
    cad = cad & vbCrLf & "Nº:  " & Text1(0).Text
    cad = cad & vbCrLf & "Fecha: " & Text1(1).Text
    cad = cad & vbCrLf & vbCrLf & " ¿Desea Eliminarlo? "
    
    'Borramos
    If MsgBox(cad, vbQuestion + vbYesNo) = vbYes Then
        Screen.MousePointer = vbHourglass
    
        NumRegElim = Data1.Recordset.AbsolutePosition

        'Abrir frame de informes para pedir datos antes de grabar en el historico
'[Monica] 08/11/2010 : no se graba en el hco de albaranes eliminados
'        cadList = ""
'        Set frmList = New frmListadoOfer
'        frmList.OpcionListado = 80
'        frmList.Show vbModal
'        Set frmList = Nothing
'        If cadList = "" Then Exit Sub
        
        If Not eliminar() Then
            CargaGrid Me.DataGrid1, Me.Data2, True
            Exit Sub
        ElseIf SituarDataTrasEliminar(Data1, NumRegElim) Then
            PonerCampos
        Else
            LimpiarCampos
            'Poner los grid sin apuntar a nada
            LimpiarDataGrids
            PonerModo 0
        End If
    End If
    Screen.MousePointer = vbDefault
    
EEliminar:
    Screen.MousePointer = vbDefault
    If Err.Number > 0 Then MuestraError Err.Number, "Eliminar Albaran", Err.Description
End Sub


Private Sub BotonEliminarLinea()
'Eliminar una linea De Mantenimiento. Tabla: slima1
Dim sql As String

    'Si no estaba modificando lineas salimos
    'Es decir, si estaba insertando linea no podemos hacer otra cosa
    If ModificaLineas = 1 Or ModificaLineas = 2 Then Exit Sub '1= Insertar, 2=Modificar

    If Data2.Recordset.EOF Then Exit Sub
        
    ModificaLineas = 3 'Eliminar
    sql = "¿Seguro que desea eliminar la línea de Albaran?     "
    sql = sql & vbCrLf & "NumLinea:  " & Data2.Recordset!numlinea & vbCrLf
    sql = sql & "Almacen:  " & Format(Data2.Recordset!codAlmac, "000")
    sql = sql & vbCrLf & "Artículo:  " & Data2.Recordset!codartic & " - " & Data2.Recordset!NomArtic
    
    If MsgBox(sql, vbQuestion + vbYesNo) = vbYes Then
        'Hay que eliminar
        NumRegElim = Data2.Recordset.AbsolutePosition
        If EliminarLinea Then
            ModificaLineas = 0
            SituarDataTrasEliminar Data2, NumRegElim
            CargaGrid2 DataGrid1, Data2
            CalcularDatosFactura
            '++monica: rollo
            PonerModo 2
            
        End If
'        CancelaADODC
    End If
    PonerFocoBtn Me.cmdRegresar

EEliminarLinea:
        Screen.MousePointer = vbDefault
        If Err.Number <> 0 Then MuestraError Err.Number, "Eliminar Lineas Mantenimientos", Err.Description
End Sub


Private Sub cmdRegresar_Click()
'Este es el boton Cabecera
Dim cad As String

    'Quitar lineas y volver a la cabecera
    If Modo = 5 Then  'modo 5: Mantenimientos Lineas
        PonerModo 2
        'BloquearTabs False
        If Not Data1.Recordset.EOF Then _
            Me.lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
        If DataGrid1.Row >= 0 Then
            DeseleccionaGrid DataGrid1
            DataGrid1.Bookmark = 1
        End If
    Else 'Se llama desde algún Prismatico de otro Form al Mantenimiento de Trabajadores
        If Data1.Recordset.EOF Then
            MsgBox "Ningún registro devuelto.", vbExclamation
            Exit Sub
        End If
        cad = Data1.Recordset.Fields(0) & "|"
        cad = cad & Data1.Recordset.Fields(1) & "|"
        RaiseEvent DatoSeleccionado(cad)
        Unload Me
    End If
End Sub


Private Sub DataGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
Dim devuelve As String
Dim cadLote As String

    On Error GoTo Error1
    
    If Not Data2.Recordset.EOF And ModificaLineas <> 1 Then '1: Insertar
'--monica
'        cadLote = "numlotes"
'        devuelve = DevuelveDesdeBDNew(cptours, NomTablaLineas, "ampliaci", "numalbar", Text1(0).Text, "N", cadLote, "numlinea", Data2.Recordset!numlinea, "N")
        devuelve = DevuelveDesdeBDNew(cPTours, NomTablaLineas, "ampliaci", "numalbar", Text1(0).Text, "T", , "numlinea", Data2.Recordset!numlinea, "N")
        'Poner descripcion de ampliacion lineas
        text2(16).Text = devuelve
        'poner el numero de lote
        text2(17).Text = cadLote
        CargarDatosArticulo (Data2.Recordset!codartic)
    End If
    If ModificaLineas = 1 Then
        text2(16).Text = ""
        text2(0).Text = ""
        text2(1).Text = ""
    End If

Error1:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub


Private Sub Form_Activate()
    Screen.MousePointer = vbDefault
    
    If PrimeraVez Then
        PrimeraVez = False
        'Viene de DblClick en frmAlmMovimArticulos y carga el form con los valores
        If hcoCodMovim <> "" And Not Data1.Recordset.EOF Then PonerCadenaBusqueda

        
        'Viene de click en VerAlbaranes en formulario de "Recepcion de Facturas compra"
        If cadSelAlbaranes <> "" And Not Data1.Recordset.EOF Then PonerCadenaBusqueda
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub Form_Load()
Dim i  As Integer
'    'Icono del formulario
'    Me.Icon = frmPpal.Icon
    
    PrimeraVez = True

    ' ICONITOS DE LA BARRA
    btnAnyadir = 5
    btnPrimero = 16
    With Me.Toolbar1
        .HotImageList = frmPpal.imgListComun_OM
        .DisabledImageList = frmPpal.imgListComun_BN
        .ImageList = frmPpal.imgListComun
        .Buttons(1).Image = 1   'Botón Buscar
        .Buttons(2).Image = 2   'Botón Todos
        .Buttons(5).Image = 3   'Insertar Nuevo
        .Buttons(6).Image = 4   'Modificar
        .Buttons(7).Image = 5   'Borrar
        .Buttons(9).Image = 15 'Mto Lineas Albaran
        .Buttons(10).Image = 15 'Mto Lineas Albaran
'--monica
'        .Buttons(11).Image = 33 'Nº Serie

'        .Buttons(12).Image = 16 'Imprimir Albaran
        
        .Buttons(13).Image = 11  'Salir
        .Buttons(btnPrimero).Image = 6  'Primero
        .Buttons(btnPrimero + 1).Image = 7 'Anterior
        .Buttons(btnPrimero + 2).Image = 8 'Siguiente
        .Buttons(btnPrimero + 3).Image = 9 'Último
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
    ' ***********************************
    'cargar IMAGES de busqueda
    For i = 0 To Me.imgBuscar.Count - 1
        Me.imgBuscar(i).Picture = frmPpal.imgListImages16.ListImages(1).Picture
    Next i
    
    Me.SSTab1.Tab = 0
    LimpiarCampos   'Limpia los campos TextBox
   
    CodTipoMov = "ALC"
    VieneDeBuscar = False

    '## A mano
     Me.FrameHco.visible = EsHistorico
    
    If Not EsHistorico Then
        NombreTabla = "scaalp"
        NomTablaLineas = "slialp" 'Tabla lineas de Albaranes
        Me.Caption = "Albaranes Proveedores"
        Ordenacion = " ORDER BY numalbar, fechaalb,codprove "
    Else
        NombreTabla = "schalp"
        NomTablaLineas = "slhalp"
        CargarTagsHco Me, "scaalp", NombreTabla
        'Estos campos solo estan en la tabla del histórico
        Text1(22).Tag = "Fecha Eliminación|F|N|||schalp|fechelim|dd/mm/yyyy|N|"
'--monica
'        Text1(23).Tag = "Trabajador Eliminación|N|N|0|9999|schalp|trabelim|0000|N|"
        Text1(24).Tag = "Incidencia elim.|T|N|||schalp|codincid||N|"
        Me.Caption = "Histórico Albaranes Proveedores"
        Ordenacion = " ORDER BY numalbar,fechaalb,codprove "
    End If
    
         
    'Vemos como esta guardado el valor del check
    chkVistaPrevia.Value = CheckValueLeer(Name)
    
    'ASignamos un SQL al DATA1
    CadenaConsulta = "Select * from " & NombreTabla
    
    If hcoCodMovim <> "" Then
    'Se llama desde Dobleclick en frmAlmMovimArticulos
        CadenaConsulta = CadenaConsulta & " WHERE numalbar='" & hcoCodMovim & "' AND fechaalb= """ & Format(hcoFechaMovim, "yyyy-mm-dd") & """"
        CadenaConsulta = CadenaConsulta & " AND codprove=" & hcoCodProve
    ElseIf cadSelAlbaranes <> "" Then
        CadenaConsulta = CadenaConsulta & " WHERE " & cadSelAlbaranes
    Else
        CadenaConsulta = CadenaConsulta & " WHERE numalbar = -1"
    End If
    Data1.ConnectionString = Conn
    Data1.RecordSource = CadenaConsulta
    Data1.Refresh
       
    'Cargar el DataGrid de lineas de Revisiones inicialmente a nada DATA2

    If hcoCodMovim = "" Then
        If DatosADevolverBusqueda = "" Then
            PonerModo 0
        Else
            PonerModo 1
            Text1(0).BackColor = vbYellow
        End If
        'Poner los grid sin apuntar a nada
    End If
   
    LimpiarDataGrids
    PrimeraVez = True

End Sub


Private Sub LimpiarCampos()
On Error Resume Next
    limpiar Me   'Metodo general: Limpia los controles TextBox
    lblIndicador.Caption = ""
    'Aqui va el especifico de cada form es
    '### a mano
    Text3(0).Text = "BASE IMP."
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub Form_Unload(Cancel As Integer)
    CheckValueGuardar Me.Name, Me.chkVistaPrevia.Value
    If Modo = 4 Or Modo = 5 Then TerminaBloquear
End Sub


Private Sub frmAlm_DatoSeleccionado(CadenaSeleccion As String)
'Form Mantenimiento de Almacenes Propios
    txtAux(0).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000") 'Cod Almacen
End Sub

Private Sub frmArt_DatoSeleccionado(CadenaSeleccion As String)
'Mantenimiento de Articulos
    txtAux(1).Text = RecuperaValor(CadenaSeleccion, 1) 'Cod Artic
    txtAux(2).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub


Private Sub frmB_Selecionado(CadenaDevuelta As String)
Dim CadB As String
Dim Aux As String
      
    If CadenaDevuelta <> "" Then
        HaDevueltoDatos = True
        Screen.MousePointer = vbHourglass
        CadB = ""
        Aux = ValorDevueltoFormGrid(Text1(0), CadenaDevuelta, 1)
        CadB = Aux
        Aux = ValorDevueltoFormGrid(Text1(1), CadenaDevuelta, 2)
        CadB = CadB & " and " & Aux
        Aux = ValorDevueltoFormGrid(Text1(4), CadenaDevuelta, 3)
        CadB = CadB & " and " & Aux
        
        CadenaConsulta = "select * from " & NombreTabla & " WHERE " & CadB & " " & Ordenacion
        PonerCadenaBusqueda
    End If
    Screen.MousePointer = vbDefault
End Sub


Private Sub frmF_Selec(vFecha As Date) 'Calendario Fechas
    Text1(CByte(imgFec(0).Tag)).Text = Format(vFecha, "dd/mm/yyyy")
End Sub


Private Sub frmFP_DatoSeleccionado(CadenaSeleccion As String)
'Form Mantenimiento de Formas de Pago
Dim indice As Byte
    indice = 12
    Text1(indice).Text = RecuperaValor(CadenaSeleccion, 1) 'Cod Forma Pago
    FormateaCampo Text1(indice)
    text2(indice).Text = RecuperaValor(CadenaSeleccion, 2) 'Nom Forma Pago
End Sub


Private Sub frmList_DatoSeleccionado(CadenaSeleccion As String)
'devuelve los datos necesarios para grabar en la tabla del historico al eliminar albaran
    cadList = ""
    cadList = DBSet(RecuperaValor(CadenaSeleccion, 1), "F") & " as fechelim,"
'--monica: no hay trabajador
'    cadList = cadList & RecuperaValor(CadenaSeleccion, 2) & " as trabelim,"
    cadList = cadList & DBSet(RecuperaValor(CadenaSeleccion, 2), "N") & " as codincid"
End Sub


'--monica
'Private Sub frmMen_DatoSeleccionado(CadenaSeleccion As String)
'Dim cant As Currency
'Dim i As Byte
'Dim cadSerie As String
'Dim nSerie As CNumSerie
'
''si llegamos aqui hemos hecho un abono y vamos a eliminar el
''nº de serie de la tabla sserie del articulo que hemos devuelto.
'
'    cant = CCur(txtAux(3).Text)
'    cant = Abs(cant)
'
'    'Para cada valor empipado actualizar la tabla sserie
'    On Error GoTo ErrorNSerie
'
'    For i = 1 To cant
'        cadSerie = RecuperaValor(CadenaSeleccion, i + 1) 'Cod Forma Pago
'        If cadSerie <> "" Then
'            Set nSerie = New CNumSerie
'            nSerie.numserie = cadSerie
'            nSerie.Articulo = RecuperaValor(CadenaSeleccion, 1)
'
'            'como vamos a devolver esos nº serie de ese articulo
'            'los eliminamos de la tabla sserie, ya no tenemos esos artículos
'            nSerie.EliminarNumSerie
'            Set nSerie = Nothing
'        End If
'    Next i
'
'ErrorNSerie:
'    If Err.Number <> 0 Then
'        MuestraError Err.Number, "Actualizar tabla Nº Series", Err.Description
'    End If
'End Sub



Private Sub frmProve_DatoSeleccionado(CadenaSeleccion As String)
'Form Mantenimiento de Proveedores
    Text1(4).Text = RecuperaValor(CadenaSeleccion, 1)  'Cod Prove
    If Modo = 1 Then 'Busqueda
        Text1(5).Text = PonerNombreDeCod(Text1(4), "proveedor", "nomprove")
    Else 'Si Insertar, recuperar datos de Tabla sprove
        PonerDatosProveedor (Text1(4).Text)
    End If
    
'
'    FormateaCampo Text1(4)
'    Text1(5).Text = RecuperaValor(CadenaSeleccion, 2)  'Descripcion
End Sub

Private Sub frmPV_DatoSeleccionado(CadenaSeleccion As String)
'Form Mantenimiento de Proveedores Varios
Dim indice As Byte

    indice = 6
    Text1(indice).Text = RecuperaValor(CadenaSeleccion, 1) 'NIF
    Text1(indice - 1).Text = RecuperaValor(CadenaSeleccion, 2) 'Nom prove
    PonerDatosProveVario (Text1(indice).Text)
End Sub

Private Sub imgBuscar_Click(Index As Integer)
Dim indice As Byte

    If Modo = 2 Or Modo = 0 Then Exit Sub
    Screen.MousePointer = vbHourglass

    Select Case Index
        Case 0 'Cod. Prove
            PonerFoco Text1(4)
            Set frmProve = New frmManProve
            frmProve.DatosADevolverBusqueda = "0|1|"
            frmProve.Show vbModal
            Set frmProve = Nothing
            indice = 4
'--monica
'        Case 1 'Realizada Por Trabajador
'            indice = 2
'            Set frmT = New frmAdmTrabajadores
'            frmT.DatosADevolverBusqueda = "0"
'            frmT.Show vbModal
'            Set frmT = Nothing
'
'        Case 2 'Cod. Postal
'            Set frmCP = New frmCPostal
'            frmCP.DatosADevolverBusqueda = "0"
'            frmCP.Show vbModal
'            Set frmCP = Nothing
'            indice = 9
'            VieneDeBuscar = True
            
        Case 3 'Forma de Pago
            indice = 12
            PonerFoco Text1(indice)
            Set frmFP = New frmManFpago
            frmFP.DatosADevolverBusqueda = "0|"
            frmFP.Show vbModal
            Set frmFP = Nothing
            
        Case 5 'NIF proveedor varios
'            Set frmPV = New frmComProveV
'            frmPV.DatosADevolverBusqueda = "0|"
'            frmPV.Show vbModal
'            Set frmPV = Nothing
'            indice = 6
    End Select
    PonerFoco Text1(indice)
    Screen.MousePointer = vbDefault
End Sub


Private Sub imgFec_Click(Index As Integer) 'Abre calendario Fechas
Dim indice As Byte

   If Modo = 2 Or Modo = 0 Then Exit Sub
   Screen.MousePointer = vbHourglass
   
   '++monica
    Dim esq As Long
    Dim dalt As Long
    Dim menu As Long
    Dim obj As Object

    Set frmF = New frmCal
    
    esq = imgFec(Index).Left
    dalt = imgFec(Index).Top
    
    Set obj = imgFec(Index).Container

    While imgFec(Index).Parent.Name <> obj.Name
        esq = esq + obj.Left
        dalt = dalt + obj.Top
        Set obj = obj.Container
    Wend
    
    menu = Me.Height - Me.ScaleHeight 'ací tinc el heigth del menú i de la toolbar

    frmF.Left = esq + imgFec(Index).Parent.Left + 30
    frmF.Top = dalt + imgFec(Index).Parent.Top + imgFec(Index).Height + menu - 40

   
   frmF.NovaData = Now
   Select Case Index
        Case 0
            indice = 1
        Case 1
            indice = 25
   End Select
   Me.imgFec(0).Tag = indice
   
   PonerFormatoFecha Text1(indice)
   If Text1(indice).Text <> "" Then frmF.NovaData = CDate(Text1(indice).Text)

   Screen.MousePointer = vbDefault
   frmF.Show vbModal
   Set frmF = Nothing
   PonerFoco Text1(indice)
End Sub


Private Sub mnBuscar_Click()
    Me.SSTab1.Tab = 0
    BotonBuscar
End Sub


Private Sub mnEliminar_Click()
    If Modo = 5 Then 'Eliminar lineas de Pedido
         BotonEliminarLinea
    Else   'Eliminar Pedido
         BotonEliminar
         Screen.MousePointer = vbDefault
    End If
End Sub


Private Sub mnLineas_Click()
    BotonMtoLineas 0, "Albaranes"
End Sub


Private Sub mnModificar_Click()
    If Modo = 5 Then 'Modificar lineas
         BotonModificarLinea
    Else   'Modificar Cabecera Albaran
        If cadSelAlbaranes = "" Then
            If Not BLOQUEADesdeFormulario(Me) Then Exit Sub
        End If
        BotonModificar
    End If
End Sub


Private Sub mnNuevo_Click()
    If Modo = 5 Then 'Añadir lineas
         BotonAnyadirLinea
    Else 'Añadir Cabecera de Albaran
         Me.SSTab1.Tab = 0
         BotonAnyadir
    End If
End Sub


Private Sub mnSalir_Click()
    Screen.MousePointer = vbDefault
    If (Modo = 5) Then 'Modo 5: Mto Lineas
        '1:Insertar linea, 2: Modificar
        If ModificaLineas = 1 Or ModificaLineas = 2 Then cmdCancelar_Click
        cmdRegresar_Click
        Exit Sub
    End If
    Unload Me
End Sub

Private Sub mnVerTodos_Click()
    BotonVerTodos
End Sub

'### A mano
'Los metodos del text tendran que estar
'Los descomentamos cuando esten puestos ya los controles
Private Sub Text1_GotFocus(Index As Integer)
    kCampo = Index
    If Index = 9 Then HaCambiadoCP = False 'CPostal
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
Dim devuelve As String
        
    If Not PerderFocoGnral(Text1(Index), Modo) Then Exit Sub
    
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub
          
    'Si queremos hacer algo ..
    Select Case Index
        Case 1, 25 'Fecha Albaran y fecha de turno
            PonerFormatoFecha Text1(Index)
                
            
        Case 4 'Cod. Proveedor
            If PonerFormatoEntero(Text1(Index)) Then
                If Modo = 1 Then 'Busqueda
                    Text1(5).Text = PonerNombreDeCod(Text1(Index), "proveedor", "nomprove")
                Else 'Si Insertar, recuperar datos de Tabla sprove
                    PonerDatosProveedor (Text1(Index).Text)
                End If
            Else
                LimpiarDatosProve
            End If
            
         Case 6 'NIF
            If Not EsDeVarios Or Modo <> 3 Then Exit Sub
            If Modo = 4 Then 'Modificar
                'si no se ha modificado el nif del cliente no hacer nada
                If Text1(6).Text = Data1.Recordset!NIFProve Then
                    Exit Sub
                End If
            End If
            PonerDatosProveVario (Text1(Index).Text)
            
        Case 12 'Forma de Pago
            If PonerFormatoEntero(Text1(Index)) Then
                text2(Index).Text = PonerNombreDeCod(Text1(Index), "sforpa", "nomforpa")
            Else
                text2(Index).Text = ""
            End If
            
        Case 13, 14 'Descuentos
            If PonerFormatoDecimal(Text1(Index), 4) Then 'Tipo 4: Decimal(4,2)
                If Modo = 4 Then CalcularDatosFactura
                If Index = 14 Then
                    Me.SSTab1.Tab = 1
                    PonerFoco Text1(15)
                End If
            Else
                If Index = 14 And Text1(Index).Text = "" Then
                    Me.SSTab1.Tab = 1
                    PonerFoco Text1(15)
                End If
            End If
            
        Case 26 ' codigo de turno
            PonerFormatoEntero Text1(Index)
    End Select
End Sub


Private Sub HacerBusqueda()
Dim CadB As String

    CadB = ObtenerBusqueda3(Me, False)
    If cadSelAlbaranes <> "" Then CadB = CadB & " AND " & cadSelAlbaranes
    If chkVistaPrevia = 1 Then
        MandaBusquedaPrevia CadB
    ElseIf CadB <> "" Then
        'Se muestran en el mismo form
        CadenaConsulta = "select * from " & NombreTabla & " WHERE " & CadB & " " & Ordenacion
        PonerCadenaBusqueda
    End If
End Sub


Private Sub MandaBusquedaPrevia(CadB As String)
'Carga el formulario frmBuscaGrid con los valores correspondientes
Dim cad As String
Dim tabla As String
Dim Titulo As String

    'Llamamos a al form
    '##A mano
    cad = ""
    cad = cad & ParaGrid(Text1(0), 20, "Nº Albaran")
    cad = cad & ParaGrid(Text1(1), 15, "Fecha Alb.")
    cad = cad & ParaGrid(Text1(4), 15, "Provedor")
    cad = cad & ParaGrid(Text1(5), 50, "Nombre Prov.")
    tabla = NombreTabla
    Titulo = "Albaranes"
           
    If cad <> "" Then
        Screen.MousePointer = vbHourglass
        Set frmB = New frmBuscaGrid
        frmB.vCampos = cad
        frmB.vTabla = tabla
        frmB.vSQL = CadB
        HaDevueltoDatos = False
        '###A mano
        frmB.vDevuelve = "0|1|2|"
        frmB.vTitulo = Titulo
        frmB.vSelElem = 0
'        frmB.vConexionGrid = conAri  'Conexión a BD: Ariges

        frmB.Show vbModal
        Set frmB = Nothing
        'Si ha puesto valores y tenemos que es formulario de busqueda entonces
        'tendremos que cerrar el form lanzando el evento
'        If HaDevueltoDatos Then
'''            If (Not Data1.Recordset.EOF) And DatosADevolverBusqueda <> "" Then _
'''                cmdRegresar_Click
'        Else   'de ha devuelto datos, es decir NO ha devuelto datos
'            PonerFoco Text1(kCampo)
        'End If
    End If
    Screen.MousePointer = vbDefault
End Sub


Private Sub PonerCadenaBusqueda()
Screen.MousePointer = vbHourglass
On Error GoTo EEPonerBusq

    Data1.RecordSource = CadenaConsulta
    Data1.Refresh
    If Data1.Recordset.RecordCount <= 0 Then
        MsgBox "No hay ningún registro en la tabla " & NombreTabla, vbInformation
        Screen.MousePointer = vbDefault
        If Modo = 1 Then
            PonerFoco Text1(0)
            Text1(0).BackColor = vbYellow
        End If
        Exit Sub
    Else
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


Private Sub PonerCamposLineas()
'Carga las Pestañas con las tablas de lineas del Trabajador seleccionado para mostrar
On Error GoTo EPonerLineas

    Screen.MousePointer = vbHourglass

    'Datos de la tabla slipre
    CargaGrid DataGrid1, Data2, True

    Screen.MousePointer = vbDefault
    Exit Sub
EPonerLineas:
    MuestraError Err.Number, "PonerCamposLineas"
    PonerModo 2
    Screen.MousePointer = vbDefault
End Sub


Private Sub PonerCampos()
On Error Resume Next

    If Data1.Recordset.EOF Then Exit Sub
    PonerCamposForma Me, Data1
    
'--monica
'    'Trabajador Albaran
'    Text2(2).Text = PonerNombreDeCod(Text1(2), conAri, "straba", "nomtraba", "codtraba")
    text2(12).Text = PonerNombreDeCod(Text1(12), "sforpa", "nomforpa")
'--monica
'    'Trabajador del Pedido
'    Text2(21).Text = PonerNombreDeCod(Text1(21), conAri, "straba", "nomtraba", "codtraba")
    
    PonerCamposLineas 'Pone los datos de las tablas de lineas de Ofertas
    
    If EsHistorico Then
        'poner datos de eliminacion
'--monica
'        Text2(23).Text = PonerNombreDeCod(Text1(23), conAri, "straba", "nomtraba", "codtraba")
        text2(24).Text = PonerNombreDeCod(Text1(24), "inciden", "nomincid", "codincid")
    End If
    
    CalcularDatosFactura 'rellenar campos pestaña de totales
     
    '-- Esto permanece para saber donde estamos
    lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
    
    PonerModoOpcionesMenu Modo
    PonerOpcionesMenu
    
    If Err.Number <> 0 Then Err.Clear
End Sub


'----------------------------------------------------------------
'----------------------------------------------------------------
'   En PONERMODO se habilitan, o no, los diverso campos del
'   formulario en funcion del modo en k vayamos a trabajar
Private Sub PonerModo(Kmodo As Byte)
Dim i As Integer, NumReg As Byte
Dim b As Boolean

    On Error GoTo EPonerModo
    
'--monica: rollo toolbar
'    'Actualiza Iconos Insertar,Modificar,Eliminar
'    ActualizarToolbarGnral Me.Toolbar1, Modo, Kmodo, btnAnyadir
    
    Modo = Kmodo
    PonerIndicador lblIndicador, Modo
    
    'Modo 2. Hay datos y estamos visualizandolos
    '=========================================
    b = (Modo = 2)
    'Ponemos visible, si es formulario de busqueda, el boton regresar cuando hay datos
    If DatosADevolverBusqueda <> "" Or hcoCodMovim <> "" Then
        cmdRegresar.visible = b
    Else
        cmdRegresar.visible = False
    End If
        
    'Poner Flechas de desplazamiento visibles
    NumReg = 1
    If Not Data1.Recordset.EOF Then
        If Data1.Recordset.RecordCount > 1 Then NumReg = 2 'Solo es para saber q hay + de 1 registro
    End If
    DesplazamientoVisible Me.Toolbar1, btnPrimero, b, NumReg
        
        
    'Bloquea los campos Text1 sino estamos modificando/Insertando Datos
    'Si estamos en Insertar además limpia los campos Text1 y bloquea la clave primaria
'    BloquearText1 Me, Modo
    b = (Modo = 1 Or Modo = 3 Or Modo = 4)
    For i = 0 To Text1.Count - 1
        BloquearTxt Text1(i), Not b
        Text1(i).Enabled = b
    Next i
    
    ' la clave primaria esta bloqueada
    BloquearTxt Text1(0), (Modo <> 1) And (Modo <> 3) '
    BloquearTxt Text1(4), (Modo <> 1) And (Modo <> 3)
    
    'La fecha de albaran es clave primaria pero dejamos modificarla
    BloquearTxt Text1(1), (Modo = 0 Or Modo = 2 Or Modo = 5)
    Text1(1).Enabled = Not (Modo = 0 Or Modo = 2 Or Modo = 5)
    
    b = (Modo <> 1)
    'Bloquear los campos de Pedido, excepto en Busqueda
    BloquearTxt Text1(3), b
    Text1(3).Enabled = Not b
    BloquearTxt Text1(20), b
    Text1(20).Enabled = Not b
    BloquearTxt Text1(21), b
    Text1(21).Enabled = Not b
    
    'datos cliente siempre bloqueados hasta que sea de varios
    If Modo = 3 Then
        EsDeVarios = False
        BloquearDatosProve (EsDeVarios)
    End If
     
    '-----  Datos Totales de Factura siempre bloqueado
    For i = 33 To 50
        BloquearTxt Text3(i), True
    Next i
    'Campo B.Imp y Imp. IVA siempre en azul
    Text3(36).BackColor = &HFFFFC0
    Text3(46).BackColor = &HFFFFC0
    Text3(47).BackColor = &HFFFFC0
    Text3(48).BackColor = &HFFFFC0
    Text3(49).BackColor = &HC0C0FF    'Tatal factura
    Text3(50).BackColor = &HC0C0FF    'Tatal factura
    '---------------------------------------------------
          
    'Si no es modo lineas Boquear los TxtAux
    For i = 0 To txtAux.Count - 1
        BloquearTxt txtAux(i), (Modo <> 5)
    Next i
    txtAux(2).Enabled = False
    
    BloquearTxt text2(16), (Modo <> 5)
    BloquearTxt text2(0), True
    BloquearTxt text2(1), True
    
    '---------------------------------------------
    b = (Modo = 3 Or Modo = 4 Or Modo = 1 Or Modo = 5) '++monica: rollo toolbar añadido modo = 5
    cmdCancelar.visible = b
    cmdAceptar.visible = b
    
    For i = 0 To Me.imgFec.Count - 1
'        Me.imgFecha(i).Enabled = b
        BloquearImgFec Me, i, Modo
    Next i
    b = (Modo = 3 Or Modo = 4 Or Modo = 1)
    
    For i = 0 To Me.imgBuscar.Count - 1
        Me.imgBuscar(i).Enabled = b
    Next i
    Me.imgBuscar(4).Enabled = (Modo = 1)
    Me.imgBuscar(0).Enabled = (Modo = 3 Or Modo = 1)
              
    'Modo Linea de Albaranes. Campo Ampliacion Linea
'--monica: rollo toolbar
'    Me.Label1(35).visible = (Modo = 5)
'    Me.Text2(16).visible = (Modo = 5)
    BloquearTxt text2(16), True
    'Modo Linea de Albaranes. Campo num_lote
    Me.Label1(2).visible = (Modo = 5)
    Me.text2(17).visible = (Modo = 5)
    BloquearTxt text2(17), True
       
    Me.chkVistaPrevia.Enabled = (Modo <= 2)
       
    'Poner el tamaño de los campos. Si es modo Busqueda el MaxLength del campo
    'debe ser mayor para adminir intervalos de busqueda.
    PonerLongCampos
    
    PonerModoOpcionesMenu (Modo) 'Activar opciones de menu según modo
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
'Comprobar que los datos de la cabecera son correctos antes de Insertar o Modificar
'la cabecera del Pedido
Dim b As Boolean

    On Error GoTo EDatosOK

    DatosOk = False
       
    b = CompForm(Me) 'Comprobar formato datos ok
    If Not b Then Exit Function
    
    DatosOk = b
EDatosOK:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Function


Private Function DatosOkLinea() As Boolean
Dim b As Boolean
Dim i As Byte
Dim cArt As CArticulo

    On Error GoTo EDatosOkLinea

    DatosOkLinea = False
    b = True

    'Comprobar que los campos requeridos tengan valor
    For i = 0 To txtAux.Count - 1
        If txtAux(i).Text = "" Then
            Screen.MousePointer = vbDefault
            MsgBox "El campo " & txtAux(i).Tag & " no puede ser nulo", vbExclamation
            b = False
            PonerFoco txtAux(i)
            Exit Function
        End If
    Next i
    
    
    'si el articulo tiene control de numero de lotes, el campo del lote será requerido
    Set cArt = New CArticulo
    If cArt.LeerDatos(txtAux(1).Text) Then
        If cArt.TieneNumLote Then
            If Trim(text2(17).Text) = "" Then
                b = False
                MsgBox "El nº de lote no puede ser nulo." & vbCrLf & vbCrLf & "El artículo tiene control de lotes.", vbExclamation
            End If
        End If
    End If
    Set cArt = Nothing
    
'    If Me.Text2(17).Locked = False Then
'        If Trim(Text2(17).Text) = "" Then
'            b = False
'            MsgBox "El nº de lote no puede ser nulo." & vbCrLf & vbCrLf & "El artículo tiene control de lotes.", vbExclamation
'        End If
'    End If
    
        
    DatosOkLinea = b
EDatosOkLinea:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Function



Private Sub Text2_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = 40 Then 'campo num_lote y Flecha hacia abajo
        If Index = 16 And text2(17).Locked Then PonerFocoBtn Me.cmdAceptar
        If Index = 17 Then PonerFocoBtn Me.cmdAceptar
    Else
        KEYdown KeyCode
    End If
End Sub

Private Sub Text2_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then 'campo Amliacion Linea y ENTER
       If Index = 16 And text2(17).Locked Then
            PonerFocoBtn Me.cmdAceptar
       ElseIf Index = 17 Then
            PonerFocoBtn Me.cmdAceptar
            
        Else
            KEYpress KeyAscii
        End If
    End If
End Sub

Private Sub Text2_LostFocus(Index As Integer)
    'quitamos los espacios en blanco
    text2(Index).Text = Trim(text2(Index).Text)
    
    If Index = 16 And (text2(Index).Locked = False) Then text2(Index).Text = UCase(text2(Index).Text)
End Sub

'++monica : rollo toolbar
Private Sub ToolAux_ButtonClick(Index As Integer, ByVal Button As MSComctlLib.Button)
Dim vWhere As String

    BotonMtoLineas 0, "Albaranes"
'-- pon el bloqueo aqui
    'If BLOQUEADesdeFormulario2(Me, Data1, 1) Then
    Select Case Button.Index
        Case 1
            mnNuevo_Click 'BotonAnyadirLineas
        Case 2
            vWhere = Replace(ObtenerWhereCP(False), NombreTabla, NomTablaLineas) & " and numlinea=" & Me.Data2.Recordset.Fields(1)
            'vWhere = ObtenerWhereCP(False) & " and numlinea=" & Me.Data2.Recordset.Fields(1)
            If BloqueaRegistro(NomTablaLineas, vWhere) Then mnModificar_Click 'BotonModificarLinea
        Case 3
            mnEliminar_Click 'BotonEliminarLinea
        Case Else
    End Select
    'End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1  'Buscar
            mnBuscar_Click
        Case 2  'Todos
            mnVerTodos_Click
        Case 5  'Nuevo
            mnNuevo_Click
        Case 6  'Modificar
            mnModificar_Click
        Case 7  'Borrar
            mnEliminar_Click
        Case 9
            ModificarProveedor
        Case 10  'Lineas
            mnLineas_Click
'--monica
'        Case 11 'Nº Series
'            BotonNSeries
        Case 13    'Salir
            mnSalir_Click
        Case btnPrimero To btnPrimero + 3 'Flechas Desplazamiento
            Desplazamiento (Button.Index - btnPrimero)
    End Select
End Sub


Private Sub PonerOpcionesMenu()
    PonerOpcionesMenuGeneral Me
End Sub


Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, Modo, cerrar
    If cerrar Then Unload Me
End Sub

    
Private Function InsertarLinea(numlinea As String) As Boolean
'Inserta un registro en la tabla de lineas de Albaranes: slialb
'OUT -> NumLinea: devuelve el Nº de linea que acaba de insertar
Dim sql As String
Dim b As Boolean
Dim vCStock As CStock
Dim vArtic As CArticulo
Dim MenError As String
Dim DentroTRANS As Boolean
Dim PrecioNue As Double

    
    InsertarLinea = False
    sql = ""
    DentroTRANS = False
    
    'Conseguir el siguiente numero de linea
    sql = ObtenerWhereCP(False)
    sql = Replace(sql, NombreTabla, NomTablaLineas)
    numlinea = SugerirCodigoSiguienteStr(NomTablaLineas, "numlinea", sql)
    Me.cmdAux(0).Tag = numlinea
    
    Set vCStock = New CStock
    If Not InicializarCStock(vCStock, "E", numlinea) Then Exit Function
    
    If DatosOkLinea() Then 'Lineas de Albaranes Proveedor
        'Inserta en tabla "slialp"
        sql = "INSERT INTO " & NomTablaLineas
        sql = sql & " (numalbar, fechaalb, codprove, numlinea, codalmac, codartic, nomartic, ampliaci, cantidad, precioar, dtoline1, dtoline2, importel) "
        sql = sql & "VALUES (" & DBSet(Text1(0).Text, "T") & ", " & DBSet(Text1(1).Text, "F") & ", " & Val(Text1(4).Text) & ", " & numlinea & ", " & Val(txtAux(0).Text) & ","
        sql = sql & DBSet(txtAux(1).Text, "T") & ", " & DBSet(txtAux(2).Text, "T") & ", " & DBSet(text2(16).Text, "T") & ", "
        sql = sql & DBSet(txtAux(3).Text, "N") & ", "
        sql = sql & DBSet(txtAux(4).Text, "N") & ", " & DBSet(txtAux(5).Text, "N") & ", "
        sql = sql & DBSet(txtAux(6).Text, "N") & ", "
        sql = sql & DBSet(txtAux(7).Text, "N") & ")" '", " & DBSet(Text2(17).Text, "T") & ");"
     Else
        Set vCStock = Nothing
        Exit Function
     End If
    
    If sql <> "" Then
        On Error GoTo eInsertarLinea
        Conn.BeginTrans
        DentroTRANS = True
        
        MenError = "Insertando lineas Albaran Compras"
        Conn.Execute sql
        
        
        '==== LAURA 20/09/2006
        'Realizar estas actualizaciones antes de modificar el stock del almacen
        MenError = "Actualizar ult. fecha compra"
        '-- Actualizar en la tabla sartic el ult precio de compra y la ult. fecha compra
        Set vArtic = New CArticulo
        vArtic.Codigo = txtAux(1).Text
        'Laura 19/12/2006: calcular precio_ult_compra con el precio con descuentos, ed. importe/cantidad, en lugar de con el precio
        'b = vArtic.ActualizarUltFechaCompra(Text1(1).Text, txtAux(4).Text)
        If CCur(txtAux(3).Text) = 0 Then
            PrecioNue = 0
        Else
            PrecioNue = Round2(CCur(txtAux(7).Text) / CCur(txtAux(3).Text), 5)
        End If
        
        b = True
        
        If CCur(txtAux(3).Text) >= 0 Then ' monica: añadida condicion de que el precio sea positivo
            b = vArtic.ActualizarUltFechaCompra(Text1(1).Text, CStr(PrecioNue))
         End If
        
        'Actualizar en la tabla sartic el precio medio ponderado
        MenError = "Actualizar precio medio ponderado"
        'Laura 19/12/2006: calcular precio_ult_compra con el precio con descuentos, ed. importe/cantidad, en lugar de con el precio
        'If b Then b = vArtic.ActualizarPrecioMedPond(CCur(txtAux(3).Text), CCur(txtAux(4).Text))
        If b Then b = vArtic.ActualizarPrecioMedPond(CCur(txtAux(3).Text), PrecioNue)
        
       
        
        Set vArtic = Nothing
        '====
        
        'en actualizar stock comprobamos si el articulo tiene control de stock
        If b Then
            MenError = "Actualizando Stocks"
            b = vCStock.ActualizarStock
        End If
        
'--monica
'        If b Then
'            'si el articulo tiene control de numero de lotes, insertar en la tabla slotes
'            If Me.Text2(17).Locked = False Then
'                'si ya existe la linea aumentamos la cantidad entrada
'                SQL = "SELECT COUNT(*) FROM slotes WHERE "
'                SQL = SQL & " codartic=" & DBSet(txtAux(1).Text, "T") & " AND numlotes=" & DBSet(Text2(17).Text, "T") & " AND fecentra=" & DBSet(Text1(1).Text, "F")
'                If RegistrosAListar(SQL) > 0 Then
'                    SQL = "UPDATE slotes SET canentra=canentra + " & DBSet(txtAux(3).Text, "N")
'                    SQL = SQL & " WHERE " & " codartic=" & DBSet(txtAux(1).Text, "T") & " AND numlotes=" & DBSet(Text2(17).Text, "T") & " AND fecentra=" & DBSet(Text1(1).Text, "F")
'                Else
'                    SQL = "INSERT INTO slotes (codartic,numlotes,fecentra,canentra,canasign) VALUES ("
'                    SQL = SQL & DBSet(txtAux(1).Text, "T") & ", " & DBSet(Text2(17).Text, "T") & ", "
'                    'fecha entrada, cantidad entrada y cantidad asignada
'                    SQL = SQL & DBSet(Text1(1).Text, "F") & "," & DBSet(txtAux(3).Text, "N") & ",0)"
'                    Conn.Execute SQL
'                End If
'            End If
'        End If
    End If
    
    Set vCStock = Nothing
    
eInsertarLinea:
    If Err.Number <> 0 Then b = False
    
    If b Then
        If DentroTRANS Then Conn.CommitTrans
        InsertarLinea = True
    Else
        If DentroTRANS Then Conn.RollbackTrans
        InsertarLinea = False
        MuestraError Err.Number, "Insertar Lineas Albaran" & vbCrLf & MenError & vbCrLf, Err.Description
    End If
End Function


Private Function ModificarLinea() As Boolean
'Modifica un registro en la tabla de lineas de Albaran: slialb
Dim sql As String, vWhere As String
Dim vCStock As CStock
Dim vArtic As CArticulo
Dim b As Boolean
Dim MenError As String
Dim dentroTRANSAC As Boolean
Dim cadNumLote As String
'--monica
'Dim cLote As CNumLote

    On Error GoTo EModificarLinea

    ModificarLinea = False
    sql = ""
    dentroTRANSAC = False
    
    Set vCStock = New CStock
    If Not InicializarCStock(vCStock, "E") Then Exit Function
    
    
    Set vArtic = New CArticulo
    If Not vArtic.LeerDatos(txtAux(1).Text) Then Exit Function
    

'    If DatosOkLinea() Then
    'sql para actualizar la linea del albaran compras
    sql = "UPDATE " & NomTablaLineas & " Set codalmac = " & txtAux(0).Text & ", codartic=" & DBSet(txtAux(1).Text, "T") & ", "
    sql = sql & "nomartic=" & DBSet(txtAux(2).Text, "T") & ", ampliaci=" & DBSet(text2(16).Text, "T", "S") & ", "
    sql = sql & "cantidad= " & DBSet(txtAux(3).Text, "N") & ", "
    sql = sql & "precioar=" & DBSet(txtAux(4).Text, "N") & ", " 'precio
    sql = sql & "dtoline1= " & DBSet(txtAux(5).Text, "N") & ", dtoline2= " & DBSet(txtAux(6).Text, "N") & ", "
    sql = sql & "importel= " & DBSet(txtAux(7).Text, "N") '& ", "
'--monica
'    SQL = SQL & "numlotes=" & DBSet(Text2(17).Text, "T", "S")
    vWhere = ObtenerWhereCP(True) & " AND numlinea=" & Data2.Recordset!numlinea
    vWhere = Replace(vWhere, NombreTabla, NomTablaLineas)
    sql = sql & vWhere

    If sql <> "" Then
        dentroTRANSAC = True
        Conn.BeginTrans
            
        MenError = "Actualizando Lineas Albaran Compras"
        Conn.Execute sql
            
            
        '==== Laura 20/09/2006, antes de actualizar el stock
        ' deshacer el precio medio ponderado y luego calcularlo otra vez con los nuevos valores
        MenError = "Recalcular precio medio ponderado del articulo."
        '-- Laura 18/12/2006: calcular precio_med_pond con el precio aplicandole el descuento, ed. importe/cantidad.
        'b = vArtic.ReestablecerPrecioMedPon(CCur(Data2.Recordset!Cantidad), CCur(Data2.Recordset!precioar))
        If CCur(DBLet(Data2.Recordset!cantidad, "N")) = 0 Then
            b = vArtic.ReestablecerPrecioMedPon(CCur(Data2.Recordset!cantidad), 0)
        Else
            b = vArtic.ReestablecerPrecioMedPon(CCur(Data2.Recordset!cantidad), CCur(Data2.Recordset!importel) / CCur(Data2.Recordset!cantidad))
        End If
        '-- Laura 18/12/2006: calcular precio_med_pond con el precio aplicandole el descuento, ed. importe/cantidad.
        'If b Then b = vArtic.ActualizarPrecioMedPond(CCur(txtAux(3).Text), CCur(txtAux(4).Text), CCur(Data2.Recordset!Cantidad))
        If CCur(txtAux(3).Text) <> 0 Then
            If b Then b = vArtic.ActualizarPrecioMedPond(CCur(txtAux(3).Text), Round2(CCur(txtAux(7).Text) / CCur(txtAux(3).Text), 4), CCur(Data2.Recordset!cantidad))
        Else
            If b Then b = vArtic.ActualizarPrecioMedPond(CCur(txtAux(3).Text), 0, CCur(Data2.Recordset!cantidad))
        End If
        
        'Actualizar ultima fecha de compra del articulo
        If b Then
            MenError = "Actualizando ult. fecha compra"
            '-- Laura 18/12/2006: actualizar precio_ult_compra con el precio aplicandole el descuento, ed. importe/cantidad.
            'b = vArtic.ActualizarUltFechaCompra(Text1(1).Text, txtAux(4).Text)
            
            'monica: 17/06/2009 si la cantidad es negativa no se actualiza fecha ni precio UC
            If CCur(txtAux(3).Text) >= 0 Then
                If CCur(txtAux(3).Text) <> 0 Then
                    b = vArtic.ActualizarUltFechaCompra(Text1(1).Text, Round2(CCur(txtAux(7).Text) / CCur(txtAux(3).Text), 4))
                Else
                    b = vArtic.ActualizarUltFechaCompra(Text1(1).Text, 0)
                End If
            End If
        End If
        '====
            
            
        'Actualizar Stocks de los articulos y movimientos
        '===================================================
        If b Then
            MenError = "Actualizando stocks y movimientos almacen"
            'si no se ha modificado el almacen reestablecemos cantidad y precio
            If CInt(Data2.Recordset!codAlmac) = CInt(txtAux(0).Text) Then
'                MenError = "Actualizando Stocks"
                b = vCStock.ModificarStock(Data2.Recordset!cantidad)
            Else
                'deshacer el movimiento para el almacen anterior y devolver stock
                b = InicializarCStock(vCStock, "S") 'movim. de salida
                If b Then b = vCStock.DevolverStock
                            
                'Insertar el movimiento para el nuevo almacen y actualizar stock
                b = InicializarCStock(vCStock, "E") 'mov. de entrada
                If b Then b = vCStock.ActualizarStock
            End If
        End If
                

'--monica
'        '=== CONTROL Nº DE LOTES DEL ARTICULO
'        '===============================================
'        If b Then
'            'comprobar si el artículo que modificamos tiene control de lotes
'            MenError = "Actualizando Nº Lote."
'            If vArtic.TieneNumLote Then
'                    'si no existe en la tabla slotes lo añadimos sino lo modificamos
'                    SQL = "SELECT COUNT(*) FROM slotes "
'                    SQL = SQL & " WHERE codartic=" & DBSet(Data2.Recordset!codArtic, "T") & " AND numlotes=" & DBSet(Text2(17).Text, "T")
'                    SQL = SQL & " AND fecentra=" & DBSet(Data2.Recordset!FechaAlb, "F")
'                    If RegistrosAListar(SQL) > 0 Then
'                        'actualizar la cantidad de entrada de la tabla slotes
'                        SQL = "UPDATE slotes SET canentra=canentra + " & DBSet(CStr(CSng(txtAux(3).Text)) - CSng(Me.Data2.Recordset!Cantidad), "N")
'                        SQL = SQL & " WHERE codartic=" & DBSet(Data2.Recordset!codArtic, "T") & " AND numlotes=" & DBSet(Data2.Recordset!numlotes, "T") & " AND fecentra=" & DBSet(Data2.Recordset!FechaAlb, "F")
'                        Conn.Execute SQL
'                    ElseIf Text2(17).Text <> "" Then
'                        'SI NO EXISTE LO INSERTAMOS
'                        SQL = "INSERT INTO slotes (codartic,numlotes,fecentra,canentra,canasign) VALUES ("
'                        SQL = SQL & DBSet(Data2.Recordset!codArtic, "T") & "," & DBSet(Text2(17).Text, "T") & "," & DBSet(Data2.Recordset!FechaAlb, "F") & ","
'                        SQL = SQL & DBSet(txtAux(3).Text, "N") & ",0)"
'                        Conn.Execute SQL
'                    End If
'
'                    'SI HEMOS MODIFICADO EL Nº DE LOTE
'                    'DESCONTAMOS LA CANTIDAD DE LA LINEA DE LA VIEJA
'                    'Y SI ES CERO LA BORRAMOS
'                    If Text2(17).Text <> CStr(DBLet(Data2.Recordset!numlotes, "T")) Then
'                        If Not IsNull(Data2.Recordset!numlotes) Then
'                            If DBLet(Data2.Recordset!numlotes, "T") <> "" Then
'                                'actualizar la cantidad de entrada de la tabla slotes
'                                SQL = "UPDATE slotes SET canentra=canentra - " & DBSet(txtAux(3).Text, "N")
'                                SQL = SQL & " WHERE codartic=" & DBSet(Data2.Recordset!codArtic, "T") & " AND numlotes=" & DBSet(Data2.Recordset!numlotes, "T") & " AND fecentra=" & DBSet(Data2.Recordset!FechaAlb, "F")
'                                Conn.Execute SQL
'                                'borrar si
'                                SQL = "DELETE FROM slotes "
'                                SQL = SQL & " WHERE codartic=" & DBSet(Data2.Recordset!codArtic, "T") & " AND numlotes=" & DBSet(Data2.Recordset!numlotes, "T") & " AND fecentra=" & DBSet(Data2.Recordset!FechaAlb, "F")
'                                SQL = SQL & " AND canentra=0"
'                                Conn.Execute SQL
'                            End If
'                        End If
'                    End If
'            End If
'        End If
                
            
        If b Then
            Conn.CommitTrans
        Else
            Conn.RollbackTrans
        End If
        ModificarLinea = b
    End If
        
    
    Set vCStock = Nothing
    Set vArtic = Nothing
    Exit Function
    
EModificarLinea:
    If dentroTRANSAC Then Conn.RollbackTrans
    If Not vArtic Is Nothing Then Set vArtic = Nothing
    If Not vCStock Is Nothing Then Set vCStock = Nothing
    ModificarLinea = False
    MuestraError Err.Number, "Modificar Lineas Albaran" & vbCrLf & MenError & vbCrLf & Err.Description
End Function





Private Sub PonerBotonCabecera(b As Boolean)
'Pone el boton de Regresar a la Cabecera si pasamos a MAntenimiento de Lineas
'o Pone los botones de Aceptar y cancelar en Insert,update o delete lineas
On Error Resume Next

    Me.cmdAceptar.visible = Not b
    Me.cmdCancelar.visible = Not b
    Me.cmdRegresar.visible = b
    Me.cmdRegresar.Caption = "Cabecera"
    If b Then
        Me.lblIndicador.Caption = "Líneas " & TituloLinea
        PonerFocoBtn Me.cmdRegresar
    End If
    'Habilitar las opciones correctas del menu segun Modo
'    PonerModoOpcionesMenu (Modo)
'    PonerOpcionesMenu 'Habilitar las opciones correctas del menu segun Nivel de Acceso
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub CargaGrid(ByRef vDataGrid As DataGrid, ByRef vData As Adodc, enlaza As Boolean)
Dim b As Boolean
Dim sql As String

On Error GoTo ECargaGRid

    b = DataGrid1.Enabled
    
    sql = MontaSQLCarga(enlaza)
    CargaGridGnral vDataGrid, vData, sql, PrimeraVez
    
    CargaGrid2 vDataGrid, vData
    
    
    b = (Modo = 5) And (ModificaLineas = 1 Or ModificaLineas = 2)
    vDataGrid.Enabled = Not b
    If Modo = 2 Then vDataGrid.Enabled = True
    PrimeraVez = False
    
    DataGrid1.ScrollBars = dbgAutomatic
    Exit Sub
    
ECargaGRid:
    If Err.Number <> 0 Then MuestraError Err.Number, "Cargando datos grid", Err.Description
End Sub


Private Sub CargaGrid2(ByRef vDataGrid As DataGrid, ByRef vData As Adodc)
Dim i As Byte
On Error GoTo ECargaGRid

    vData.Refresh

    vDataGrid.Columns(0).visible = False
    vDataGrid.Columns(1).visible = False
    vDataGrid.Columns(2).visible = False
    vDataGrid.Columns(3).visible = False
    
    Select Case vDataGrid.Name
        Case "DataGrid1" 'Cod. Almacen
                i = 4
                vDataGrid.Columns(i).Caption = "Alm."
                vDataGrid.Columns(i).Width = 500
                vDataGrid.Columns(i).NumberFormat = "000"
                i = i + 1
                vDataGrid.Columns(i).Caption = "Articulo"
                vDataGrid.Columns(i).Width = 1700
                i = i + 1
                vDataGrid.Columns(i).Caption = "Desc. Artículo"
                vDataGrid.Columns(i).Width = 3400
                
                i = i + 1
                vDataGrid.Columns(i).visible = False
                i = i + 1
                vDataGrid.Columns(i).Caption = "Cantidad"
                vDataGrid.Columns(i).Width = 850
                vDataGrid.Columns(i).Alignment = dbgRight
                vDataGrid.Columns(i).NumberFormat = FormatoImporte
                
                i = i + 1
                vDataGrid.Columns(i).Caption = "Precio"
                vDataGrid.Columns(i).Width = 1200
                vDataGrid.Columns(i).Alignment = dbgRight
                vDataGrid.Columns(i).NumberFormat = FormatoDec10d5
                
                i = i + 1
                vDataGrid.Columns(i).Caption = "Dto. 1"
                vDataGrid.Columns(i).Width = 600
                vDataGrid.Columns(i).Alignment = dbgRight
                vDataGrid.Columns(i).NumberFormat = FormatoPorcen 'FormatoDescuento
                
                i = i + 1
                vDataGrid.Columns(i).Caption = "Dto. 2"
                vDataGrid.Columns(i).Width = 600
                vDataGrid.Columns(i).Alignment = dbgRight
                vDataGrid.Columns(i).NumberFormat = FormatoPorcen 'FormatoDescuento
                
                i = i + 1
                vDataGrid.Columns(i).Caption = "Importe Línea"
                vDataGrid.Columns(i).Width = 1500
                vDataGrid.Columns(i).Alignment = dbgRight
                vDataGrid.Columns(i).NumberFormat = FormatoImporte
'--monica
'                i = i + 1
'                vDataGrid.Columns(i).visible = False
    End Select

    For i = 0 To vDataGrid.Columns.Count - 1
        vDataGrid.Columns(i).Locked = True
        vDataGrid.Columns(i).AllowSizing = False
    Next i
    Exit Sub
ECargaGRid:
    If Err.Number <> 0 Then MuestraError Err.Number, "Cargando datos grid", Err.Description
End Sub


'Esta funcion sustituye a LlamaLineas
Private Sub CargaTxtAux(visible As Boolean, limpiar As Boolean)
'IN: visible: si es true ponerlos visibles en la posición adecuada
'    limpiar: si es true vaciar los txtAux
Dim alto As Single
Dim i As Byte

    If Not visible Then
        'Fijamos el alto (ponerlo en la parte inferior del form)
        For i = 0 To txtAux.Count - 1 'TextBox
            txtAux(i).Top = 290
            txtAux(i).visible = visible
        Next i
        cmdAux(0).visible = visible
        cmdAux(1).visible = visible
    Else
        If limpiar Then 'Vaciar los textBox (Vamos a Insertar)
            DeseleccionaGrid DataGrid1
            For i = 0 To txtAux.Count - 1
                txtAux(i).Text = ""
                BloquearTxt txtAux(i), False
                
                txtAux(i).Enabled = True
            Next i
        Else 'Vamos a modificar
            For i = 0 To txtAux.Count - 1
                If i < 3 Then 'campos anteriores a ampliacion linea (ampliaci)
                    txtAux(i).Text = DataGrid1.Columns(i + 4).Text
                Else
                    txtAux(i).Text = DataGrid1.Columns(i + 5).Text
                End If
                txtAux(i).Locked = False
                
            Next i
        End If
               
        ' la descripcion de articulo ha de estar siempre not enabled
        txtAux(2).Enabled = False
               
        'El campo Importe es calculado y lo bloqueamos.
'        BloquearTxt txtAux(7), True
    
    
        '#Laura 15/11/2006
        'no se puede modificar el almacen y el articulo pq no elimina bien de smoval
        'y no reestablece stock si se cambia el articulo (REVISAR!!!)
'        BloquearTxt txtAux(0), (ModificaLineas = 2) 'codalmac
        BloquearTxt txtAux(1), (ModificaLineas = 2) 'codartic
        txtAux(1).Enabled = Not (ModificaLineas = 2)
'        Me.cmdAux(0).Enabled = (ModificaLineas <> 2)
        Me.cmdAux(1).Enabled = (ModificaLineas <> 2)
    
        '#
    
    
    

        'Fijamos altura(Height) y posición Top
        '-------------------------------
        alto = ObtenerAlto(DataGrid1, 20)
        
        For i = 0 To txtAux.Count - 1
            txtAux(i).Top = alto
            txtAux(i).Height = DataGrid1.RowHeight
        Next i
        cmdAux(0).Top = alto
        cmdAux(1).Top = alto
        cmdAux(0).Height = DataGrid1.RowHeight
        cmdAux(1).Height = DataGrid1.RowHeight
        
        'Fijamos anchura y posicion Left
        '--------------------------------
        'Cod. Almac
        txtAux(0).Left = DataGrid1.Left + 330
        txtAux(0).Width = DataGrid1.Columns(4).Width - 160
        cmdAux(0).Left = txtAux(0).Left + txtAux(0).Width - 40
        'Cod Artic
        txtAux(1).Left = cmdAux(0).Left + cmdAux(0).Width + 20
        txtAux(1).Width = DataGrid1.Columns(5).Width - 160
        cmdAux(1).Left = txtAux(1).Left + txtAux(1).Width - 50
        'Nom Artic
        txtAux(2).Left = cmdAux(1).Left + cmdAux(1).Width
        txtAux(2).Width = DataGrid1.Columns(6).Width - 10
        'Cantidad
        txtAux(3).Left = txtAux(2).Left + txtAux(2).Width + 10
        txtAux(3).Width = DataGrid1.Columns(8).Width - 10
        'Precio, Dto1, Dto2, Precio
        For i = 4 To txtAux.Count - 1
            txtAux(i).Left = txtAux(i - 1).Left + txtAux(i - 1).Width + 10
            txtAux(i).Width = DataGrid1.Columns(i + 5).Width - 10
        Next i
        
        'Los ponemos Visibles o No
        '--------------------------
        For i = 0 To txtAux.Count - 1
            txtAux(i).visible = visible
        Next i
        cmdAux(0).visible = visible
        cmdAux(1).visible = visible
    End If
End Sub


Private Sub txtAux_GotFocus(Index As Integer)
Dim cadkey As Integer

    cadkey = ObtenerCadKey(kCampo, Index)
    kCampo = Index
    PonerFoco txtAux(Index)
    ConseguirFocoLin txtAux(Index)
End Sub

Private Sub Txtaux_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'Avanzar/Retroceder los campos con las flechas de desplazamiento del teclado.
    If Not (Index = 0 And KeyCode = 38) Then KEYdown KeyCode
End Sub


Private Sub txtAux_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub txtAux_LostFocus(Index As Integer)
Dim devuelve As String
Dim TipoDto As Byte
Dim b As Boolean
Dim bLotes As Boolean

    If Not PerderFocoGnralLineas(txtAux(Index), ModificaLineas) Then Exit Sub
    
    Select Case Index
        Case 0 'Cod Almacen
            'Comprobar que existe el almacen
            devuelve = PonerAlmacen(txtAux(Index).Text)
            txtAux(Index).Text = devuelve
            If devuelve = "" Then PonerFoco txtAux(Index)
    
        Case 1 'Cod. Articulo
            If txtAux(1).Text = "" Then
                txtAux(2).Text = ""
                Exit Sub
            End If
            If txtAux(0).Text = "" Then
                MsgBox "Debe seleccionar un almacen.", vbInformation
                PonerFoco txtAux(0)
                Exit Sub
            End If
            
            If PonerArticulo(txtAux(1), txtAux(2), txtAux(0).Text, CodTipoMov, ModificaLineas, , bLotes) Then
                BloquearTxt text2(17), Not bLotes
                
                CargarDatosArticulo txtAux(1).Text
                
                b = (Me.ActiveControl.Name = "txtAux")
                If b Then b = (Me.ActiveControl.Index = 0)
                
                If Not b Then
'                    If txtAux(2).Locked Then PonerFoco txtAux(3)
                Else
                    PonerFoco txtAux(0)
                End If
            Else
                PonerFoco txtAux(Index)
            End If
            
        Case 2 'Desc. Articulo
            If txtAux(Index).Locked = False Then txtAux(Index).Text = UCase(txtAux(Index).Text)
            
        Case 3 'Cantidad
            If PonerFormatoDecimal(txtAux(Index), 1) Then  'Tipo 1: Decimal(12,2)
                If (Modo = 5 And ModificaLineas = 1) Then 'Modo Insertar en Mto Lineas
                    'Obtener el precio correspondiente y los descuentos
                    ObtenerPrecioCompra
                End If
            End If

        Case 4 'Precio
            PonerFormatoDecimal txtAux(Index), 7  'Tipo 7: Decimal(10,5)
            
        Case 5, 6 'Descuentos
            PonerFormatoDecimal txtAux(Index), 4 'Tipo 4: Decimal(4,2)
        Case 7 'Importe Linea
            PonerFormatoDecimal txtAux(Index), 1 'Tipo 3: Decimal(12,2)
    End Select
    If (Index = 3 Or Index = 4 Or Index = 5 Or Index = 6) Then
'        If Trim(TxtAux(3).Text) = "" Or Trim(TxtAux(4).Text) = "" Then Exit Sub
'        If Trim(TxtAux(6).Text) = "" Or Trim(TxtAux(7).Text) = "" Then Exit Sub
        If txtAux(1).Text = "" Then Exit Sub
        TipoDto = DevuelveDesdeBDNew(cPTours, "proveedor", "tipodtos", "codprove", Text1(4).Text, "N")
        txtAux(7).Text = CalcularImporteProv(txtAux(3).Text, txtAux(4).Text, txtAux(5).Text, txtAux(6).Text, TipoDto, 0)
        PonerFormatoDecimal txtAux(7), 1
    Else
        If Index = 7 Then
            PonerFormatoDecimal txtAux(7), 1
        End If
    End If
End Sub



Private Sub ObtenerPrecioCompra()
'Dim vPrecio As CPreciosCom
Dim cad As String
'
    On Error GoTo EPrecios
'
'    Set vPrecio = New CPreciosCom
'    If vPrecio.Leer(txtAux(1).Text, Text1(4).Text) Then
'        If vPrecio.ComprobarCantidad(CInt(txtAux(3).Text)) Then
'            txtAux(4).Text = vPrecio.ObtenerPrecio(Text1(1).Text)
'            txtAux(5).Text = vPrecio.Descuento1
'            txtAux(6).Text = vPrecio.Descuento2
'        Else
'            PonerFoco txtAux(3)
'            Exit Sub
'        End If
'    Else
        'Obtener el ult. precio de compra de ese articulo (sartic)
        cad = DevuelveDesdeBDNew(cPTours, "sartic", "ultpreci", "codartic", txtAux(1).Text, "T")
        If cad <> "" Then
            txtAux(4).Text = cad
            txtAux(5).Text = "0"
            txtAux(6).Text = "0"
        End If
'    End If
    PonerFormatoDecimal txtAux(4), 7
    PonerFormatoDecimal txtAux(5), 4
    PonerFormatoDecimal txtAux(6), 4
'
'    Set vPrecio = Nothing
    
EPrecios:
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub BotonMtoLineas(numTab As Integer, cad As String)
        Me.SSTab1.Tab = numTab
        TituloLinea = cad
        ModificaLineas = 0
        PonerModo 5
'--monica: rollo toolbar
'        PonerBotonCabecera True
End Sub


Private Function eliminar() As Boolean
Dim vWhere As String
Dim b As Boolean
Dim sql As String

    On Error GoTo FinEliminar

        Conn.BeginTrans
        vWhere = " " & ObtenerWhereCP(False)
                
        
        'Reestablecer el stock en la tabla salmac a partir de todas las lineas del albaran
        'Eliminar los movimientos de smoval
        b = ReestablecerStock(vWhere)
        

        
        
        If b Then
            'Pasar los datos al historico de albaranes de compra y borrarlos de albaranes
            'scaalp --> schalp
            'slialp --> slhalp
            b = ActualizarElTraspaso("", vWhere, CodTipoMov, cadList)
            
'--monica
'            'Eliminar los numeros de serie del albaran sino estan vendidos a ningun cliente
'            If b Then
'                SQL = "DELETE FROM sserie WHERE numalbpr=" & DBSet(Data1.Recordset!numalbar, "T")
'                SQL = SQL & " AND fechacom=" & DBSet(Data1.Recordset!FechaAlb, "F")
'                SQL = SQL & " AND codprove=" & Data1.Recordset!codProve
'                SQL = SQL & " AND (isnull(numfactu) and isnull(numalbar))"
'                Conn.Execute SQL
'            End If
            
        End If
        
FinEliminar:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Eliminar Albaran Compras", Err.Description
        b = False
    End If
    If Not b Then
        Conn.RollbackTrans
        eliminar = False
    Else
        Conn.CommitTrans
        eliminar = True
    End If
End Function


Private Sub LimpiarDataGrids()
'Pone los Grids sin datos, apuntando a ningún registro
On Error Resume Next
    CargaGrid DataGrid1, Data2, False
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub PosicionarData()
Dim Indicador As String
Dim vWhere As String

    If Not Data1.Recordset.EOF Then
        'Hay datos en el Data1 bien porque se ha hecho VerTodos o una Busqueda
         vWhere = "(" & ObtenerWhereCP(False) & ")"
         vWhere = Replace(vWhere, NombreTabla & ".", "")
         If SituarDataMULTI(Data1, vWhere, Indicador) Then
             PonerModo 2
             lblIndicador.Caption = Indicador
        Else
             LimpiarCampos
             'Poner los grid sin apuntar a nada
             LimpiarDataGrids
             PonerModo 0
         End If
    Else
        'El Data esta vacio, desde el modo de inicio se pulsa Insertar
        CadenaConsulta = "Select * from " & NombreTabla & ObtenerWhereCP(True) & Ordenacion
        PonerCadenaBusqueda
    End If
End Sub


Private Function ObtenerWhereCP(conW As Boolean) As String
Dim sql As String
On Error Resume Next
    
    sql = ""
    If conW Then sql = " WHERE "
    sql = sql & NombreTabla & ".numalbar= " & DBSet(Text1(0).Text, "T") & " and " & NombreTabla & ".fechaalb='" & Format(Text1(1).Text, FormatoFecha)
    sql = sql & "' and " & NombreTabla & ".codprove=" & Val(Text1(4).Text)
    
    ObtenerWhereCP = sql
End Function


Private Function MontaSQLCarga(enlaza As Boolean) As String
'--------------------------------------------------------------------
' MontaSQlCarga:
'   Basándose en la información proporcionada por el vector de campos
'   crea un SQl para ejecutar una consulta sobre la base de datos que los
'   devuelva.
' Si ENLAZA -> Enlaza con el data1
'           -> Si no lo cargamos sin enlazar a ningun campo
'--------------------------------------------------------------------
Dim sql As String
    
    sql = "SELECT numalbar,fechaalb,codprove, numlinea, codalmac, codartic, nomartic, ampliaci, cantidad, precioar, dtoline1, dtoline2, importel " ', numlotes "
    sql = sql & " FROM " & NomTablaLineas
    If enlaza Then
        sql = sql & " " & ObtenerWhereCP(True)
    Else
        sql = sql & " WHERE numalbar = -1"
    End If
    sql = sql & " Order by numalbar, fechaalb, codprove, numlinea"
    sql = Replace(sql, NombreTabla, NomTablaLineas)
    MontaSQLCarga = sql
End Function


Private Sub PonerModoOpcionesMenu(Modo As Byte)
'Activas unas Opciones de Menu y Toolbar según el modo en que estemos
Dim b As Boolean, bAux As Boolean
Dim i As Integer

   
        b = (Modo = 2) And hcoCodMovim = "" '--monica: rollo toolbar --> Or (Modo = 5 And ModificaLineas = 0)
        'Insertar
        Toolbar1.Buttons(5).Enabled = (b Or Modo = 0) And (cadSelAlbaranes = "" Or (cadSelAlbaranes <> "" And Modo = 5)) And Not EsHistorico
        Me.mnNuevo.Enabled = (b Or Modo = 0) And (cadSelAlbaranes = "" Or (cadSelAlbaranes <> "" And Modo = 5)) And Not EsHistorico
        'Modificar
        Toolbar1.Buttons(6).Enabled = b And Not EsHistorico
        Me.mnModificar.Enabled = b And Not EsHistorico
        'eliminar
        Toolbar1.Buttons(7).Enabled = b And cadSelAlbaranes = "" And Not EsHistorico
        Me.mnEliminar.Enabled = b And cadSelAlbaranes = "" And Not EsHistorico
            
        b = (Modo = 2) And Not EsHistorico And (hcoCodMovim = "")
'--monica: rollo toolbar
'        'Mantenimiento lineas
'        Toolbar1.Buttons(10).Enabled = b
'        Me.mnLineas.Enabled = b
        Toolbar1.Buttons(11).Enabled = b And (hcoCodMovim = "")
        
        b = (Modo >= 3) Or Modo = 1
        'Buscar
        Toolbar1.Buttons(1).Enabled = (Not b) And (hcoCodMovim = "")
        Me.mnBuscar.Enabled = (Not b) And (hcoCodMovim = "")
        'Ver Todos
        Toolbar1.Buttons(2).Enabled = (Not b) And (hcoCodMovim = "")
        Me.mnVerTodos.Enabled = (Not b) And (hcoCodMovim = "")
        
    '++monica: rollo toolbar
    ' *** si n'hi han llínies que tenen grids (en o sense tab) ***
    'b = (Modo = 3 Or Modo = 4 Or Modo = 2) And Not EsHistorico
    '++monica:15102008 cambiado por esto
    b = (Modo = 2) And Not EsHistorico And hcoCodMovim = ""
    For i = 0 To ToolAux.Count - 1
        ToolAux(i).Buttons(1).Enabled = b
        If b Then bAux = (b And Me.Data2.Recordset.RecordCount > 0)
        ToolAux(i).Buttons(2).Enabled = bAux
        ToolAux(i).Buttons(3).Enabled = bAux
    Next i
        
End Sub



Private Function InsertarAlbaran(vSQL As String) As Boolean
Dim MenError As String
Dim devuelve As String
Dim bol As Boolean

    On Error GoTo EInsertarOferta
    
    bol = False
    
    'Aqui empieza transaccion
    Conn.BeginTrans
    MenError = "Error al insertar en la tabla Cabecera de Albaranes (" & NombreTabla & ")."
    Conn.Execute vSQL, , adCmdText
    
    'Actualizar los datos del proveedor si es de varios
    If EsDeVarios Then
        MenError = "Modificando datos proveedor varios."
        bol = ActualizarProveVarios(Text1(4).Text, Text1(6).Text)
    End If
    
    
    'Actualizar el campo fecha ult.compra(fechamov) en la tabla proveedores (sprove)
    devuelve = DevuelveDesdeBDNew(cPTours, "proveedor", "fechamov", "codprove", Text1(4).Text, "N")
    If (devuelve = "") Then devuelve = "01/01/1900"
    If CDate(Text1(1).Text) > CDate(devuelve) Then
        vSQL = "UPDATE proveedor SET fechamov=" & DBSet(Text1(1).Text, "F")
        vSQL = vSQL & " WHERE codprove=" & Text1(4).Text
        Conn.Execute vSQL, , adCmdText
    End If
    bol = True
    
    
EInsertarOferta:
    If Err.Number <> 0 Then
        MenError = "Insertando Albaran." & vbCrLf & "----------------------------" & vbCrLf & MenError
        MuestraError Err.Number, MenError, Err.Description
        bol = False
    End If
    If bol Then
        Conn.CommitTrans
        InsertarAlbaran = True
    Else
        Conn.RollbackTrans
        InsertarAlbaran = False
    End If
End Function


Private Sub LimpiarDatosProve()
Dim i As Byte

    For i = 4 To 14
        Text1(i).Text = ""
    Next i
End Sub
    

Private Function EliminarLinea() As Boolean
Dim vCStock As CStock
'--monica
'Dim cLote As CNumLote
Dim cArt As CArticulo
Dim sql As String
Dim b As Boolean

    EliminarLinea = False
    
    
    'Inicilizar la clase para Actualizar los stocks
    Set vCStock = New CStock
    If Not InicializarCStock(vCStock, "S") Then Exit Function
    
    '==== Laura: 20/09/2006
    'Inicializar la clase para actualizar precio medio ponderado del Articulo
    Set cArt = New CArticulo
    If Not cArt.LeerDatos(vCStock.codartic) Then Exit Function
    '====
    
    On Error GoTo EEliminarLinea
    Conn.BeginTrans
    
'--monica
'    'Eliminar las lineas de la tabla "sserie", insertadas para la linea del albaran a eliminar
'    SQL = " WHERE  numalbpr= " & DBSet(Text1(0).Text, "T") & " and fechacom='" & Format(Text1(1).Text, FormatoFecha)
'    SQL = SQL & "' and codprove=" & Val(Text1(4).Text) & " AND numline2=" & Data2.Recordset!numlinea
'    Conn.Execute "Delete from sserie " & SQL
    
    'Construir la SQL para eliminar la linea de la tabla "slialb"
    sql = ObtenerWhereCP(True) & " and numlinea=" & Data2.Recordset!numlinea
    sql = Replace(sql, NombreTabla, NomTablaLineas)
    sql = "Delete from " & NomTablaLineas & sql
    Conn.Execute sql 'Eliminar linea
    
    '==== Laura: 20/09/2006
    'reestablecer el precio medio ponderado,
    'debe calcularse antes de reestablecer el stock
    '-- Laura 19/12/2006: calcular el precio medio ponderado con precio con los descuentos ( importe/cantidad)
    'cArt.ReestablecerPrecioMedPon vCStock.Cantidad, CCur(Data2.Recordset!precioar)
    cArt.ReestablecerPrecioMedPon vCStock.cantidad, Round2(CCur(Data2.Recordset!importel) / CCur(Data2.Recordset!cantidad), 4)
    Set cArt = Nothing
    '====
    
    b = vCStock.DevolverStock
    Set vCStock = Nothing
    
'--monica
'    'Si el articulo tiene control de lotes eliminar la cantidad eliminada
'    'si la linea se queda con cero borrarla.
'    If b Then
'        If Not IsNull(Data2.Recordset!numlotes) Then
'            Set cLote = New CNumLote
'            If cLote.LeerDatos(CStr(Data2.Recordset!codArtic), CStr(Data2.Recordset!numlotes), CStr(Data2.Recordset!FechaAlb)) Then
'                b = cLote.Eliminar(CSng(Data2.Recordset!Cantidad))
'
'            End If
'            Set cLote = Nothing
'        End If
'    End If
'
    

EEliminarLinea:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Eliminar Linea Albaran " & vbCrLf & Err.Description
        b = False
    End If
    If b Then
        Conn.CommitTrans
        EliminarLinea = True
    Else
        Conn.RollbackTrans
         EliminarLinea = False
    End If
End Function


Private Function InicializarCStock(ByRef vCStock As CStock, TipoM As String, Optional numlinea As String) As Boolean
'On Error Resume Next
On Error Resume Next

    vCStock.TipoMov = TipoM 'Movimiento de Entrada o Salida
    vCStock.DetaMov = CodTipoMov '"ALC=Albaran de Compra"
    vCStock.Fechamov = Text1(1).Text
    vCStock.Trabajador = CLng(Text1(4).Text) 'En smoval guardamos el Proveedor
    vCStock.Documento = Text1(0).Text
    
    If ModificaLineas = 1 Or (ModificaLineas = 2 And TipoM = "E") Then '1=Insertar, 2=Modificar
        vCStock.codartic = txtAux(1).Text
        vCStock.codAlmac = CInt(txtAux(0).Text)
        vCStock.cantidad = CSng(ComprobarCero(txtAux(3).Text))
        vCStock.Importe = CCur(ComprobarCero(txtAux(7).Text))
    Else
        vCStock.codartic = Data2.Recordset!codartic
        vCStock.codAlmac = CInt(Data2.Recordset!codAlmac)
        vCStock.cantidad = CSng(Data2.Recordset!cantidad)
        vCStock.Importe = CCur(Data2.Recordset!importel)
    End If
    If ModificaLineas = 1 Then
         vCStock.LineaDocu = CInt(ComprobarCero(numlinea))
    Else
        vCStock.LineaDocu = CInt(Data2.Recordset!numlinea)
    End If
    If Err.Number <> 0 Then
        MsgBox "No se han podido inicializar la clase para actualizar Stock", vbExclamation
        InicializarCStock = False
    Else
        InicializarCStock = True
    End If
End Function


Private Function ReestablecerStock(cadSel As String) As Boolean
Dim vCStock As CStock
Dim cArt As CArticulo
'--monica
'Dim cLote As CNumLote
Dim b As Boolean
Dim RS As ADODB.Recordset
Dim sql As String

    On Error GoTo ERestablecer
    
    sql = "SELECT * FROM " & NomTablaLineas & " WHERE " & Replace(cadSel, NombreTabla, NomTablaLineas)
    sql = sql & " ORDER BY numlinea desc "
    
    Set RS = New ADODB.Recordset
    RS.Open sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    b = True
    While (Not RS.EOF) And b
        'para cada linea de albaran deshacemos movimientos y precios medios ponderados
        Set vCStock = New CStock
           If InicializarCStock(vCStock, "S", RS!numlinea) Then
                'estos valores hay q leerlos del RS y no del data2
                 vCStock.codartic = RS!codartic
                 vCStock.codAlmac = CInt(RS!codAlmac)
                 vCStock.cantidad = CSng(RS!cantidad)
                 vCStock.Importe = CCur(RS!importel)
                 vCStock.LineaDocu = RS!numlinea
           
                '==== Laura 20/09/2006
                'antes de actualizar el stock reestablecer el precio medio ponderado del articulo
                Set cArt = New CArticulo
                If cArt.LeerDatos(vCStock.codartic) Then
                    'Laura 19/12/2006: Calcular precio medio pond. con precio con los descuentos (importe/cantidad)
                    'If Not cArt.ReestablecerPrecioMedPon(CCur(vCStock.Cantidad), CCur(RS!precioar)) Then b = False
                    If Not cArt.ReestablecerPrecioMedPon(CCur(vCStock.cantidad), Round2(vCStock.Importe / vCStock.cantidad, 4)) Then b = False
'--monica
'                    'Si el articulo tiene control de lotes eliminar la cantidad eliminada
'                    'si la linea se queda con cero borrarla.
'                    If b Then
'                        If cArt.TieneNumLote Then
'                            Set cLote = New CNumLote
'                            If cLote.LeerDatos(cArt.Codigo, CStr(DBLet(RS!numlotes, "T")), CStr(RS!FechaAlb)) Then
'                                b = cLote.Eliminar(vCStock.Cantidad)
'
'                            End If
'                            Set cLote = Nothing
'                        End If
'                    End If
                End If
                Set cArt = Nothing
                '====
                
                
                'Actualiza el stock en salmac y borra de smoval
                'Para cada linea de albaran reestablecer el stock. Como era Mov. de Entrada
                'en Almacen ahora lo tiene que borrar(S).
                If b Then
                    If Not vCStock.DevolverStock() Then b = False
                End If
           Else
               b = False
           End If
'           Data2.Recordset.MoveNext
           Set vCStock = Nothing
    
        RS.MoveNext
    Wend
    RS.Close
    Set RS = Nothing
    
    
    '#### ANTES DEL 20/09/2006
    
'    b = True
    
'    If Not Data2.Recordset.EOF Then
''       Data2.Refresh
'       Data2.Recordset.MoveFirst
'
'       'Para cada linea de albaran reestablecer el stock. Como era Mov. de Entrada
'       'en Almacen ahora lo tiene que borrar(S).
'       While (Not Data2.Recordset.EOF) And b
'           Set vCStock = New CStock
'           If InicializarCStock(vCStock, "S", Data2.Recordset!numlinea) Then
'                '==== Laura 20/09/2006
'                'antes de actualizar el stock reestablecer el precio medio ponderado del articulo
'                Set cArt = New CArticulo
'                If cArt.LeerDatos(vCStock.codArtic) Then
'                    If Not cArt.ReestablecerPrecioMedPon(CCur(vCStock.Cantidad), CCur(Data2.Recordset!precioar)) Then b = False
'                End If
'                Set cArt = Nothing
'
'               'Actualiza el stock en salmac y borra de smoval
'                If b Then
'                    If Not vCStock.DevolverStock() Then b = False
'                End If
'           Else
'               b = False
'           End If
'           Data2.Recordset.MoveNext
'           Set vCStock = Nothing
'       Wend
'    End If

ERestablecer:
    If Err.Number <> 0 Then b = False
    If Not b Then
        ReestablecerStock = False
        MuestraError Err.Number, "Reestablecer stock.", Err.Description
    Else
        ReestablecerStock = True
    End If
End Function




Private Function ReestablecerUltFecCompra() As Boolean
Dim cArt As CArticulo
Dim sql As String
Dim b As Boolean

    On Error GoTo ERestCompra
    
    b = True
    
'    select distinct codartic from slialp
'where numalbar=2100045 and fechaalb='2006-09-15' and codprove=21
    
    
'    If Not Data2.Recordset.EOF Then
'       Data2.Refresh
'       Data2.Recordset.MoveFirst
'
'       'Para cada articulo del albaran reestablecer la fecha ultima compra
'       'y el precio ultima compra
'
'       While (Not Data2.Recordset.EOF) And b
'           Set vCStock = New CStock
'           If InicializarCStock(vCStock, "S", Data2.Recordset!numlinea) Then
'               'Actualiza el stock en salmac y borra de smoval
'               If Not vCStock.DevolverStock() Then b = False
'           Else
'               b = False
'           End If
'           Data2.Recordset.MoveNext
'           Set vCStock = Nothing
'       Wend
'    End If
    
    
    
    
    ReestablecerUltFecCompra = b
    
    
ERestCompra:
    
'    If Not b Then
        ReestablecerUltFecCompra = False
'    Else
'        ReestablecerUltFecCompra = True
'    End If
End Function





'Private Function ReestablecerPrecioMedPon() As Boolean
''reestablecer el valor del precio medio ponderado
''Dim vCStock As CStock
'Dim b As Boolean
'
'    On Error GoTo EResPMP
'
'    b = True
''    If Not Data2.Recordset.EOF Then
''       Data2.Refresh
''       Data2.Recordset.MoveFirst
''
''       'Para cada linea de albaran reestablecer el stock. Como era Mov. de Entrada
''       'en Almacen ahora lo tiene que borrar(S).
''       While (Not Data2.Recordset.EOF) And b
''           Set vCStock = New CStock
''           If InicializarCStock(vCStock, "S", Data2.Recordset!numlinea) Then
''               'Actualiza el stock en salmac y borra de smoval
''               If Not vCStock.DevolverStock() Then b = False
''           Else
''               b = False
''           End If
''           Data2.Recordset.MoveNext
''           Set vCStock = Nothing
''       Wend
''    End If
'    ReestablecerPrecioMedPon = b
'    Exit Function
'
'EResPMP:
''    If Not b Then
'        ReestablecerPrecioMedPon = False
''    Else
''        ReestablecerPrecioMedPon = True
''    End If
'End Function



Private Sub InsertarCabecera()
Dim sql As String

    sql = CadenaInsertarDesdeForm(Me)
    If sql <> "" Then
        If InsertarAlbaran(sql) Then
'                            PosicionarData
            CadenaConsulta = "Select * from " & NombreTabla & ObtenerWhereCP(True) & Ordenacion
            PonerCadenaBusqueda
            'Ponerse en Modo Insertar Lineas
            BotonMtoLineas 1, "Albaranes"
            BotonAnyadirLinea
        End If
    End If
    Me.SSTab1.Tab = 0
End Sub

'--monica
'Private Sub BotonNSeries()
'Dim cadWHERE As String, SQL As String
'Dim RSLineas As ADODB.Recordset
'
'    ModificaLineas = 4
'
'    cadWHERE = " WHERE numalbar=" & DBSet(Text1(0).Text, "T")
'    cadWHERE = cadWHERE & " and fechaalb=" & DBSet(Text1(1).Text, "F")
'    cadWHERE = cadWHERE & " and slialp.codprove=" & Text1(4).Text
'
'    'Seleccionamos aquellas lineas de albaran que tienen Nº de Serie
'    SQL = "SELECT numlinea, slialp.codartic, sum(cantidad) as cantidad "
'    SQL = SQL & " FROM slialp INNER JOIN sartic on slialp.codartic=sartic.codartic "
'    SQL = SQL & cadWHERE & " And nseriesn = 1 "
'    SQL = SQL & " GROUP BY numlinea,codartic ORDER BY Codartic "
'
'    Set RSLineas = New ADODB.Recordset
'    RSLineas.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
'
'    If Not RSLineas.EOF Then
'        'Abre el formulario de pedir nº serie al comprarlos
'        'pero mostrando los nº de serie ya asignados para poder modificarlos
'        PedirNSeries RSLineas
'    Else
'        MsgBox "No hay ninguna linea de Articulo con Control de Nº Serie", vbInformation
'    End If
'    RSLineas.Close
'    Set RSLineas = Nothing
'    ModificaLineas = 0
'    DescargarDatosTMPNumSeries ("tmpnseries")
'End Sub
'
'
'Private Sub PedirNSeries(ByRef RS As ADODB.Recordset)
'Dim RSseries As ADODB.Recordset
'Dim SQL As String
'Dim linea As Integer
'
'    On Error GoTo EPedirNSeries
'
'        'Inicializo la tabla temporal de los num.serie
'        PedirNSeriesGnral RS, False
'
'        RS.MoveFirst
'        While Not RS.EOF
'            linea = 0
'            'Cargar los Nº de serie asignados al albaran
'            SQL = "SELECT numserie, codartic FROM sserie "
'            SQL = SQL & " WHERE numalbpr=" & DBSet(Text1(0).Text, "T")
'            SQL = SQL & " and fechacom='" & Format(Text1(1).Text, FormatoFecha) & "'"
'            SQL = SQL & " and codprove=" & Text1(4).Text & " and numline2=" & RS!numlinea
'            SQL = SQL & " ORDER BY codartic "
'
'            Set RSseries = New ADODB.Recordset
'            RSseries.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
'            While Not RSseries.EOF
'                linea = linea + 1
'                SQL = "UPDATE tmpnseries SET numserie=" & DBSet(RSseries!numserie, "T")
'                SQL = SQL & " WHERE codusu=" & vUsu.Codigo & " and codartic=" & DBSet(RS!codArtic, "T")
'                SQL = SQL & " and numlinealb=" & RS!numlinea
'                SQL = SQL & " and numlinea=" & linea
'                Conn.Execute SQL
'                RSseries.MoveNext
'            Wend
'            RS.MoveNext
'        Wend
'        RSseries.Close
'        Set RSseries = Nothing
'
'        SQL = "select count(*) from tmpnseries Where codusu=" & vUsu.Codigo
'        If RegistrosAListar(SQL) > 0 Then
'            Set frmNSerie = New frmRepCargarNSerie
'            frmNSerie.DeVentas = False 'Se llama desde Alb. de compras
'            frmNSerie.NumAlb = Text1(0).Text
'            frmNSerie.Show vbModal
'            Set frmNSerie = Nothing
'            espera 0.2
'        Else
'            MsgBox "No hay nº de serie asignados a ese albaran", vbInformation
'        End If
'
'EPedirNSeries:
'    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
'End Sub


'--monica
'Private Sub CargarNumSeries()
''Insertar un registro en la tabla "sserie" por cada uno de los
''Nº de Serie introducidos en la Tabla Temporal o actualizalo
'Dim SQL As String
'Dim b As Boolean
'
'    On Error GoTo ECargar
'    Conn.BeginTrans
'
'    'Borrar todos los Nº de Serie asignados a ese albaran de compra
'    'y que no tienen asignado ya un albaran de venta
'    SQL = "DELETE FROM sserie "
'    SQL = SQL & " WHERE codprove=" & Val(Text1(4).Text) & " and numalbpr=" & DBSet(Text1(0).Text, "T")
'    SQL = SQL & " and fechacom='" & Format(Text1(1).Text, FormatoFecha) & "'"
'    SQL = SQL & " and (isnull(numalbar) and isnull(numfactu))"
'    Conn.Execute SQL
'
'    'Si algun Nº serie tenia asignado albaran venta y no lo pude borrar entonces limpiamos
'    'los campos del albaran de compra
'    SQL = "UPDATE sserie SET codprove=" & ValorNulo & ", numalbpr=" & ValorNulo & ", fechacom="
'    SQL = SQL & ValorNulo & ", numline2=" & ValorNulo
'    SQL = SQL & " WHERE codprove=" & Val(Text1(4).Text) & " and numalbpr=" & DBSet(Text1(0).Text, "T")
'    SQL = SQL & " and fechacom='" & Format(Text1(1).Text, FormatoFecha) & "'"
'    Conn.Execute SQL
'
'    b = InsertarNumSeriesDeTMP
'
'
'ECargar:
'    If Err.Number <> 0 Then b = False
'    If b Then
'        Conn.CommitTrans
'    Else
'        Conn.RollbackTrans
'    End If
'
'End Sub
'
'
'Private Function InsertarNumSeriesDeTMP() As Boolean
''Inserta en la tabla sserie todos los nº de serie q se han cargado en la temporal
'Dim SQL As String
'Dim numalbar As String
'Dim b As Boolean
'Dim RStmp As ADODB.Recordset
'Dim nSerie As CNumSerie
'
'    On Error GoTo EInsertarNSeries
'
'    'Inicializamos el objeto nº de serie con los valores comunes a todos
'    Set nSerie = New CNumSerie
'    nSerie.Proveedor = CInt(Text1(4).Text)
'    nSerie.NumAlbProve = Text1(0).Text
'    nSerie.fechacom = Text1(1).Text
'
'
'    'Recuperar los Nº Serie de ese articulo cargados en la Temporal
'    'Seleccionar los nº de serie cargados en la temporal: tmpnseries
'    SQL = "SELECT * FROM tmpnseries WHERE codusu=" & vUsu.Codigo
'    SQL = SQL & " ORDER BY codartic, numlinealb, numlinea "
'
'    Set RStmp = New ADODB.Recordset
'    RStmp.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
'
'    b = True
'    While Not RStmp.EOF And b
'        nSerie.numserie = RStmp!numserie
'        nSerie.Articulo = RStmp!codArtic
'        nSerie.NumLinAlbPr = RStmp!numlinealb
'
'        'obtenemos los dias de garantia del articulo
'        SQL = DevuelveDesdeBDNew(conAri, "sartic", "garantia", "codartic", RStmp!codArtic, "T")
'        'fin garantia= fecha albaran + dias de garantia
'        nSerie.FinGarantia = CStr(CDate(Text1(1).Text) + CInt(ComprobarCero(SQL)))
'
'        'Comprobar si existe en la tabla sserie ese nº de serie
'        numalbar = "numalbpr" 'Nº albaran de Venta prove
'        SQL = DevuelveDesdeBDNew(conAri, "sserie", "numserie", "numserie", RStmp!numserie, "T", numalbar, "codartic", RStmp!codArtic, "T")
'        If SQL <> "" Then
'            If numalbar = "" Then 'ya existe el nº serie y actualizamos ya que no esta asignado a ningun albaran
'                b = nSerie.ActualizarNumSerie(False)
'            End If
'        Else
'            b = nSerie.InsertarNumSerie
'        End If
'
''        b = InsertarNSerie(RStmp!NumSerie, RStmp!codArtic, RStmp!NumLinealb)
'        RStmp.MoveNext
'    Wend
'    RStmp.Close
'    Set RStmp = Nothing
'
'    Set nSerie = Nothing
'
'EInsertarNSeries:
'    If Err.Number <> 0 Then b = False
'    If Not b Then
'        InsertarNumSeriesDeTMP = False
'    Else
'        InsertarNumSeriesDeTMP = True
'    End If
'End Function
'



Private Sub PonerDatosProveedor(CodProve As String, Optional NIFProve As String)
Dim vProve As CProveedor
Dim observaciones As String
    
    On Error GoTo EPonerDatos
    
    If CodProve = "" Then
        LimpiarDatosProve
        Exit Sub
    End If

    Set vProve = New CProveedor
    'si se ha modificado el proveedor volver a cargar los datos
    If vProve.Existe(CodProve) Then
        If vProve.LeerDatos(CodProve) Then
           
            EsDeVarios = vProve.DeVarios
            BloquearDatosProve (EsDeVarios)
        
            If Modo = 4 And EsDeVarios Then 'Modificar
                'si no se ha modificado el proveedor no hacer nada
                If CLng(Text1(4).Text) = CLng(Data1.Recordset!CodProve) Then
                    Set vProve = Nothing
                    Exit Sub
                End If
            End If
        
            Text1(4).Text = vProve.Codigo
            FormateaCampo Text1(4)
            If (Modo = 3) Or (Modo = 4) Then
                Text1(5).Text = vProve.Nombre  'Nom prove
                Text1(8).Text = vProve.Domicilio
                Text1(9).Text = vProve.CPostal
                Text1(10).Text = vProve.Poblacion
                Text1(11).Text = vProve.Provincia
                Text1(6).Text = vProve.NIF
                Text1(7).Text = DBLet(vProve.TfnoAdmon, "T")
            End If
            
            If Modo = 3 Then 'insertar
                Text1(12).Text = vProve.ForPago
                text2(12).Text = PonerNombreDeCod(Text1(12), "sforpa", "nomforpa")
                Text1(13).Text = Format(vProve.DtoPPago, FormatoPorcen)
                Text1(14).Text = Format(vProve.DtoGnral, FormatoPorcen)
            End If

            observaciones = DBLet(vProve.observaciones)
            If observaciones <> "" Then
                MsgBox observaciones, vbInformation, "Observaciones del proveedor"
            End If
        End If
    Else
        LimpiarDatosProve
    End If
    Set vProve = Nothing

EPonerDatos:
    If Err.Number <> 0 Then MuestraError Err.Number, "Poner Datos Proveedor", Err.Description
End Sub


Private Sub PonerDatosProveVario(NIFProve As String)
'Poner el los campos Text el valor del cliente
Dim vProve As CProveedor
Dim b As Boolean
   
    If NIFProve = "" Then Exit Sub
   
    Set vProve = New CProveedor
    b = vProve.LeerDatosProveVario(NIFProve)
    If b Then
        Text1(5).Text = vProve.Nombre   'Nom proveedor
        Text1(8).Text = vProve.Domicilio
        Text1(9).Text = vProve.CPostal
        Text1(10).Text = vProve.Poblacion
        Text1(11).Text = vProve.Provincia
        Text1(7).Text = DBLet(vProve.TfnoAdmon, "T")
    End If
    Set vProve = Nothing
End Sub



Private Sub BloquearDatosProve(bol As Boolean)
Dim i As Byte

    'bloquear/desbloquear campos de datos segun sea de varios o no
    If Modo <> 5 Then
        Me.imgBuscar(5).visible = bol 'NIF
        Me.imgBuscar(5).Enabled = bol 'NIF
        Me.imgBuscar(2).Enabled = bol 'poblacion
        
        For i = 5 To 11 'si no es de varios no se pueden modificar los datos
            BloquearTxt Text1(i), Not bol
        Next i
    End If
End Sub


Private Function ActualizarProveVarios(Prove As String, NIF As String) As Boolean
Dim vProve As CProveedor

    On Error GoTo EActualizarCV

    ActualizarProveVarios = False
    
    Set vProve = New CProveedor
    If EsProveedorVarios(Prove) Then
        vProve.NIF = NIF
        vProve.Nombre = Text1(5).Text
        vProve.Domicilio = Text1(8).Text
        vProve.CPostal = Text1(9).Text
        vProve.Poblacion = Text1(10).Text
        vProve.Provincia = Text1(11).Text
        vProve.TfnoAdmon = Text1(7).Text
        vProve.ActualizarProveV (NIF)
    End If
    Set vProve = Nothing
    
    ActualizarProveVarios = True
    
EActualizarCV:
    If Err.Number <> 0 Then
        ActualizarProveVarios = False
    Else
        ActualizarProveVarios = True
    End If
End Function



Private Sub CalcularDatosFactura()
Dim i As Byte
Dim cadwhere As String
Dim vFactu As CFacturaCom

    'Limpiar en el form los datos calculados de la factura
    'y volvemos a recalcular
    For i = 33 To 50
         Text3(i).Text = ""
    Next i
    
    cadwhere = ObtenerWhereCP(False)
    
    Set vFactu = New CFacturaCom
    vFactu.DtoPPago = CCur(ComprobarCero(Text1(13).Text))
    vFactu.DtoGnral = CCur(ComprobarCero(Text1(14).Text))
    If vFactu.CalcularDatosFactura(cadwhere, NombreTabla, NomTablaLineas) Then
        Text3(33).Text = vFactu.BrutoFac
        Text3(34).Text = vFactu.ImpPPago
        Text3(35).Text = vFactu.ImpGnral
        Text3(36).Text = vFactu.BaseImp
        Text3(37).Text = QuitarCero(vFactu.TipoIVA1)
        Text3(38).Text = QuitarCero(vFactu.TipoIVA2)
        Text3(39).Text = QuitarCero(vFactu.TipoIVA3)
        Text3(40).Text = vFactu.PorceIVA1
        Text3(41).Text = vFactu.PorceIVA2
        Text3(42).Text = vFactu.PorceIVA3
        Text3(43).Text = vFactu.BaseIVA1
        Text3(44).Text = vFactu.BaseIVA2
        Text3(45).Text = vFactu.BaseIVA3
        Text3(46).Text = vFactu.ImpIVA1
        Text3(47).Text = vFactu.ImpIVA2
        Text3(48).Text = vFactu.ImpIVA3
        Text3(49).Text = vFactu.TotalFac
        Text3(50).Text = vFactu.BaseImp
        
        FormatoDatosTotales
    Else
        MuestraError Err.Number, "Calculando Totales", Err.Description
    End If
    Set vFactu = Nothing
End Sub



Private Sub FormatoDatosTotales()
Dim i As Byte

    For i = 33 To 36
        If i = 34 Or i = 35 Then Text3(i).Text = QuitarCero(Text3(i).Text)
        Text3(i).Text = Format(Text3(i).Text, FormatoImporte)
    Next i
    
    'Desglose B.Imponible por IVA
    For i = 43 To 45
        If Text3(i).Text <> "" Then
             If CSng(Text3(i).Text) = 0 And Text3(i - 6).Text = "" Then
                Text3(i).Text = QuitarCero(Text3(i).Text)
                Text3(i - 3).Text = QuitarCero(Text3(i - 3).Text)
                Text3(i - 6).Text = QuitarCero(Text3(i - 6).Text)
                Text3(i + 3).Text = QuitarCero(Text3(i + 3).Text)
            Else
                Text3(i).Text = Format(Text3(i).Text, FormatoImporte)
                Text3(i - 3) = Format(Text3(i - 3).Text, FormatoPorcen)
    '            Text3(i - 6) = Format(Text3(i - 6).Text, "000")
                Text3(i + 3).Text = Format(Text3(i + 3).Text, FormatoImporte)
            End If
        End If
    Next i
    
    'Formatear el total de Factura
    Text3(49).Text = Format(Text3(49).Text, FormatoImporte)
    Text3(50).Text = Format(Text3(50).Text, FormatoImporte)
End Sub




'Private Sub ComprobarNumSeries(numlinea As String)
''Comprobamos para una linea de Albaran si el articulo tiene control de nº de serie
''y procedemos
'Dim SQL As String
'Dim cadW As String
'Dim RSLineas As ADODB.Recordset
''Dim Mostrar As Boolean 'Indica si vamos a pedir num series o a mostrarlos
''Dim cant As Integer 'cantidad que vamos a insertar
'
'
'    'si la cantidad es >0 pedimos nº serie articulos comprados
'    'si la cantidad es <0 mostramos los nº serie para devolver (ABONOS)
'
'    SQL = DevuelveDesdeBDNew(conAri, "sartic", "nseriesn", "codartic", txtAux(1).Text, "T")
'
'    If SQL = "1" Then 'Si el Articulo tiene control de nº de serie
''        If Modo = 5 Then
''            If ModificaLineas = 1 Then 'INSERTAR linea
''                If CCur(txtAux(3).Text) > 0 Then 'cantidad linea
''                    Mostrar = False
''                    cant = CSng(txtAux(3).Text)
''                ElseIf CCur(txtAux(3).Text) < 0 Then 'cantidad linea
''                    'Es un ABONO
''                    'cantidad es < 0 (es un abono, devolvemos el articulo comprado)
''                    Mostrar = True
''                End If
''
''            ElseIf ModificaLineas = 2 Then 'MODIFICAR linea
''                'comprobar que la cantidad introducida se ha modificado
''                If CSng(txtAux(3).Text) <> CSng(Data2.Recordset!Cantidad) Then
''                    cant = CSng(txtAux(3).Text) - CSng(Data2.Recordset!Cantidad)
''                    If cant > 0 Then 'añadir nuevos num serie
''                        Mostrar = False
''                    ElseIf cant < 0 Then 'mostrar num serie y quitar el que toca
''                        Mostrar = True
''                    End If
''                Else
''                    Exit Sub
''                End If
''            End If
''        End If
'
'
'
'        If CCur(txtAux(3).Text) > 0 Then 'cantidad
''        If Mostrar = False Then
'                SQL = "El Articulo tiene control de Nº de Serie." & vbCrLf & vbCrLf
'                SQL = SQL & "Introduzca los Nº de Serie"
'                If ModificaLineas = 2 Then
'                    SQL = SQL & " que se han añadido"
'                End If
'                MsgBox SQL & "." & vbCrLf, vbInformation
'                'Cargar la tabla temporal con tantas filas como cantidad de Articulo
'                'Para introducir el Nº de Serie
'                DescargarDatosTMPNumSeries "tmpnseries"
'                CargarDatosTMPNumSeries "tmpnseries", txtAux(1).Text, CInt(txtAux(3).Text), numlinea
'                'Visualizar en pantalla el Grid, y rellenar los Nº Serie
'                ModificaLineas = 0
'                Set frmNSerie = New frmRepCargarNSerie
'                frmNSerie.DeVentas = False
'                frmNSerie.NumAlb = ""
'                frmNSerie.Show vbModal
'                Set frmNSerie = Nothing
'
'        Else   'cantidad es < 0 (es un ABONO, devolvemos el articulo comprado)
'
'            'Comprobar que efectivamente estan en tabla sserie los NºSerie del Articulo
'            ' y que no esten asignados ya a otro albaran de venta
'            SQL = " select distinct count(numserie) from sserie "
'            cadW = " WHERE codartic=" & DBSet(txtAux(1).Text, "T")
'            cadW = cadW & " and codprove=" & Text1(4).Text
'            cadW = cadW & " and (numalbar='' or isnull(numalbar))"
'            SQL = SQL & cadW
'
'            If RegistrosAListar(SQL) > 0 Then 'Hay Nº de Serie para elegir
'                'mostrar los nº de serie de ese proveedor que no esten vendidos y selecccionar
'                'el que vamos a devolver
'                'Seleccionamos aquellas lineas de albaran que tienen Nº de Serie
'                SQL = "SELECT codartic, sum(cantidad) as cantidad, numlinea "
'                SQL = SQL & " FROM " & NomTablaLineas
'
'                cadW = " WHERE numalbar=" & DBSet(Text1(0).Text, "T") & " and "
'                cadW = cadW & " fechaalb=" & DBSet(Text1(1).Text, "F")
'                cadW = cadW & " and codprove= " & Text1(4).Text & " and numlinea=" & numlinea
'
'                SQL = SQL & cadW
'                SQL = SQL & " GROUP BY codartic ORDER BY Codartic "
'
'                Set RSLineas = New ADODB.Recordset
'                RSLineas.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
'
'                MostrarNSeries RSLineas
'                RSLineas.Close
'                Set RSLineas = Nothing
'            End If
'        End If
'    End If
'End Sub


'--monica
'Private Sub MostrarNSeries(ByRef RSLineas As ADODB.Recordset)
''Si los Nº de serie se introdujeron en ALBARAN COMPRAS se muestran
''los Nº de serie de los articulos comprados y se seleccionamos
''los que vamos a devolver (Para ABONOS)
'Dim SQL As String
'Dim Campos As String
'On Error GoTo EMostrarNSeries
'
'    SQL = MostrarNSeriesGnral(RSLineas, Campos)
'    SQL = SQL & " and sserie.codprove=" & Text1(4).Text
'
'    Set frmMen = New frmMensajes
'    frmMen.cadWHERE = SQL
'    frmMen.OpcionMensaje = 4 'Nº Series Articulo
'    frmMen.vCampos = Campos
'    frmMen.Show vbModal
'    Set frmMen = Nothing
'
'EMostrarNSeries:
'    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
'End Sub


Private Function ModificarCabAlbaran() As Boolean
Dim b As Boolean
Dim MenError As String

    On Error GoTo EModificaAlb

    Conn.BeginTrans

    MenError = "Modificando fecha de albaran en tablas relacionadas."
    b = ComprobarCambioFecha
                
    If b Then
        MenError = "Modificando proveedor en las tablas relacionadas."
        b = ComprobarCambioProveedor
        If b Then
            MenError = "Modificando el albaran (scaalb)."
            b = ModificaDesdeFormulario1(Me, 1)
        End If
        If b Then
            'Actualizar los datos del Proveedor si es de varios
            MenError = "Actualizando proveedor de varios."
            b = ActualizarProveVarios(Text1(4).Text, Text1(6).Text)
        End If
    End If

EModificaAlb:
    If Err.Number <> 0 Then b = False
    If b Then
        Conn.CommitTrans
    Else
        Conn.RollbackTrans
        MsgBox "Error Modificando el albaran." & vbCrLf & MenError, vbExclamation
    End If
    ModificarCabAlbaran = b
    espera 0.2
End Function




'Private Function ArticuloTieneMargen() As Boolean
'Dim cad As String
'
'    'Comprobar que el artículo tiene margen comercial
'    cad = DevuelveDesdeBDNew(conAri, "sartic", "margecom", "codartic", txtAux(1).Text, "T")
'    If cad = "" Then
'        cad = "NO SE HAN PODIDO ACTUALIZAR LOS PRECIOS." & vbCrLf
'        cad = cad & "El artículo no tiene margen comercial para calcular nuevos precios."
'        MsgBox cad, vbExclamation
'        ArticuloTieneMargen = False
'        Exit Function
'    End If
'
'
'    'comprobar que las tarifas tienen margen comercial
'    cad = "SELECT count(*)"
'    cad = cad & " FROM slista INNER JOIN starif ON slista.codlista = starif.codlista "
'    cad = cad & " WHERE slista.codartic=" & DBSet(txtAux(1).Text, "T") & " AND  isnull(margecom)"
'    If RegistrosAListar(cad) > 0 Then
'        cad = "NO SE HAN PODIDO ACTUALIZAR LOS PRECIOS." & vbCrLf
'        cad = cad & "El artículo tiene tarifas sin %PVP necesario para calcular nuevos precios."
'        MsgBox cad, vbExclamation
'        ArticuloTieneMargen = False
'        Exit Function
'    End If
'
'    ArticuloTieneMargen = True
'
'End Function


Private Function ComprobarCambioProveedor() As Boolean
'Comprueba si se ha modificado el campo proveedor de la cabecera.
'Ya que es clave primaria y se deberan cambiar tambien la fecha
'en tablas sliap y smoval
Dim RS As ADODB.Recordset
Dim sql As String
Dim Izquierda As String, Derecha As String
Dim llis As Collection
Dim i As Integer
Dim b As Boolean


    If Data1.Recordset.EOF Then Exit Function

    
    If (CDate(Text1(4).Text) <> CDate(Data1.Recordset!CodProve)) Then
    'si ha modificado el codigo de proveedor
        On Error GoTo EComprobar
        
        'seleccionar todas las lineas de ese albaran para actualizar la fecha (slialp)
        sql = "SELECT * FROM " & NomTablaLineas & " WHERE numalbar=" & DBSet(Data1.Recordset!numalbar, "T")
        sql = sql & " AND fechaalb=" & DBSet(Data1.Recordset!fechaalb, "F")
        sql = sql & " AND codprove=" & Data1.Recordset!CodProve
        
        Set RS = New ADODB.Recordset
        RS.Open sql, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
        
        Set llis = New Collection
            
        'Nos guardamos todas las lineas con la modificacion de la fecha para
        'volverlas a insertar
        BACKUP_TablaIzquierda RS, Izquierda
        
        While Not RS.EOF
            BACKUP_Tabla RS, Derecha, "codprove", CStr(Text1(4).Text)
            llis.Add Derecha
            RS.MoveNext
        Wend
        
        RS.Close
        Set RS = Nothing
        
        'Eliminamos las lineas que tenia ese albaran (slialp) para volverlas a insertar con el proveedor nuevo
        sql = "DELETE from slialp WHERE numalbar = " & DBSet(Data1.Recordset!numalbar, "T")
        sql = sql & " AND fechaalb=" & DBSet(Data1.Recordset!fechaalb, "F")
        sql = sql & " AND codprove=" & Data1.Recordset!CodProve
        Conn.Execute sql
        
        'Actualizamos la fecha en la cabecera (scaalp)
        sql = "UPDATE scaalp SET codprove = " & DBSet(Text1(4).Text, "N")
        sql = sql & " WHERE numalbar = " & DBSet(Data1.Recordset!numalbar, "T")
        sql = sql & " AND fechaalb=" & DBSet(Data1.Recordset!fechaalb, "F")
        sql = sql & " AND codprove=" & Data1.Recordset!CodProve
        Conn.Execute sql
        
        'Actualizamos la fecha en la tabla smoval
        sql = "UPDATE smoval SET codigope=" & DBSet(Text1(4).Text, "N")
        sql = sql & " WHERE document = " & DBSet(Data1.Recordset!numalbar, "T")
        sql = sql & " AND fechamov=" & DBSet(Data1.Recordset!fechaalb, "F")
        sql = sql & " AND codigope=" & Data1.Recordset!CodProve
        sql = sql & " AND detamovi='" & CodTipoMov & "'"
        Conn.Execute sql
        
        'Volvemos a insertar las lineas con el codigo de proveedor correcto (slialp)
        sql = ""
        For i = 1 To llis.Count
            If (i Mod 10) = 0 Then
                sql = sql & CStr(llis(i)) & ","
                sql = Mid(sql, 1, Len(sql) - 1) 'quitamos ultima coma
                sql = "INSERT INTO " & NomTablaLineas & " " & Izquierda & " VALUES " & sql & ";"
                Conn.Execute sql
                sql = ""
            Else
                sql = sql & CStr(llis(i)) & ","
            End If
        Next i
        
        If sql <> "" Then
            sql = Mid(sql, 1, Len(sql) - 1) 'quitamos ultima coma
            sql = "INSERT INTO " & NomTablaLineas & " " & Izquierda & " VALUES " & sql & ";"
            Conn.Execute sql
            sql = ""
        End If
        Set llis = Nothing
    End If
    b = True
    
EComprobar:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "El campo proveedor no se ha podido modificar", Err.Description
        b = False
    End If
    If b Then
        ComprobarCambioProveedor = True
    Else
        ComprobarCambioProveedor = False
    End If
End Function


Private Sub CargarDatosArticulo(codartic As String)
Dim RS As ADODB.Recordset
Dim sql As String
        
    On Error GoTo eCargarDatosArticulo
        
    If Trim(codartic) <> "" Then
'        SQL = "select nomfamia, nomunida from sartic, sfamia, sunida "
'        SQL = SQL & " where sartic.codartic = " & DBSet(codArtic, "T")
'        SQL = SQL & " and sartic.codfamia = sfamia.codfamia and sartic.codunida = sunida.codunida"
        
        sql = "select nomfamia  from sartic, sfamia "
        sql = sql & " where sartic.codartic = " & DBSet(codartic, "T")
        sql = sql & " and sartic.codfamia = sfamia.codfamia "
        
        
        Set RS = New ADODB.Recordset
        RS.Open sql, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
        
        text2(0).Text = ""
        text2(1).Text = ""
        If Not RS.EOF Then
            text2(0).Text = DBLet(RS.Fields(0).Value, "T")
            text2(1).Text = ""
        End If
        
        Set RS = Nothing
    End If
    Exit Sub
    
eCargarDatosArticulo:
    MuestraError Err.Number, "Error Cargar Datos Articulos"
End Sub



Private Sub ModificarProveedor()
Dim OK As Boolean
    OK = True
    If EsHistorico Then
        OK = False
    Else
            If Modo = 2 Then
                If Data1.Recordset Is Nothing Then
                    OK = False
                Else
                    If Data1.Recordset.EOF Then
                       OK = False
                    Else
                        'data1.Recordset!esdevarios
                        If EsDeVarios Then
                            MsgBox "Proveedor de VARIOS", vbExclamation
                            OK = False
                        End If
                    End If
                End If
            Else
                OK = False
            End If
    End If
    If OK Then
        If vSesion.Nivel > 1 Then
            MsgBox "usuario sin permiso", vbExclamation
            OK = False
        End If
    End If
    If Not OK Then Exit Sub
    

    
    CadenaDesdeOtroForm = ""
    frmListadoOfer.OpcionListado = 1
    frmListadoOfer.Show vbModal
    If CadenaDesdeOtroForm = "" Then Exit Sub
    'Si es el mismo no hago nada
'[Monica]15/11/2010 puede que me hayan cambiado los datos del proveedor
'    If (CLng(CadenaDesdeOtroForm)) = CLng(Text1(4).Text) Then
'        MsgBox "Mismo proveedor", vbExclamation
'        Exit Sub
'    End If
    

        
        
     Screen.MousePointer = vbHourglass
    'pedimos el nuevo proveedor
    Set miRsAux = New ADODB.Recordset
    Conn.BeginTrans
    OK = HacerUpdatesCodProve(CLng(CadenaDesdeOtroForm))
    If OK Then
        'Situamos
        Conn.CommitTrans
    Else
        Conn.RollbackTrans
    End If
    If OK Then
        'Sitauremos
        CadenaDesdeOtroForm = " numalbar = " & DBSet(Text1(0).Text, "T") & " AND fechaalb = " & DBSet(Text1(1).Text, "F") & " AND codprove = " & CadenaDesdeOtroForm
        CadenaConsulta = "select * from " & NombreTabla & " WHERE " & CadenaDesdeOtroForm & " " & Ordenacion
        PonerCadenaBusqueda
        
    End If
    Screen.MousePointer = vbDefault
    
    
    
End Sub




Private Function HacerUpdatesCodProve(NuevoProve As Long) As Boolean
Dim CadenaLineas As String
Dim J As Integer
Dim sql As String
Dim vPr As CProveedor
        
        On Error GoTo EHacerUpdatesCodProve
        HacerUpdatesCodProve = False
        
        Set vPr = New CProveedor
        If Not vPr.LeerDatos(CStr(NuevoProve)) Then
            Set vPr = Nothing
            Exit Function
        End If
        
            
        
        
        sql = "Select "
        sql = sql & "fechaalb,numalbar,codartic,ampliaci,nomartic,numlinea,codalmac,cantidad,precioar,"
        sql = sql & "dtoline1,dtoline2,importel,codprove"
        sql = sql & " FROM slialp"
        sql = sql & " WHERE numalbar = " & DBSet(Text1(0).Text, "T") & " AND fechaalb = " & DBSet(Text1(1).Text, "F") & " AND codprove = " & Text1(4).Text
        miRsAux.Open sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        CadenaLineas = ""
        While Not miRsAux.EOF
            sql = ", ('" & Format(miRsAux!fechaalb, FormatoFecha) & "'"
            'Texto
            For J = 1 To 4
                If IsNull(miRsAux.Fields(J)) Then
                    sql = sql & ",NULL"
                Else
                    sql = sql & ",'" & DevNombreSQL(miRsAux.Fields(J)) & "'"
                End If
            Next J
            'Numero
                        'Texto
            For J = 5 To 11
                If IsNull(miRsAux.Fields(J)) Then
                    sql = sql & ",NULL"
                Else
                    sql = sql & "," & TransformaComasPuntos(CStr(miRsAux.Fields(J)))
                End If
            Next J
            'Nuevo proveedor
            sql = sql & "," & NuevoProve
            CadenaLineas = CadenaLineas & sql & ")"

            'Sig
            miRsAux.MoveNext
        Wend
        miRsAux.Close
        'Borramos las lineas
        If CadenaLineas <> "" Then
                
            sql = "DELETE FROM slialp WHERE numalbar = " & DBSet(Text1(0).Text, "T") & " AND fechaalb = " & DBSet(Text1(1).Text, "F") & " AND codprove = " & Text1(4).Text
            Conn.Execute sql
                    
            sql = "INSERT INTO slialp ("
            sql = sql & "fechaalb,numalbar,codartic,ampliaci,nomartic,numlinea,codalmac,cantidad,precioar,"
            sql = sql & "dtoline1,dtoline2,importel,codprove"
            'Quito la primara coma
            CadenaLineas = Mid(CadenaLineas, 2)
            sql = sql & ") VALUES " & CadenaLineas
            CadenaLineas = sql
        
        End If
        
        'ACtualizamos
        'Busco los datos del proveedor
        sql = "UPDATE scaalp SET codprove = " & NuevoProve
        'Resto de datos del proveedor: nomprove,domprove,codpobla,pobprove,proprove,nifprove,telprove
        sql = sql & ", nomprove=" & DBSet(vPr.Nombre, "T")
        sql = sql & ",domprove=" & DBSet(vPr.Domicilio, "T")
        sql = sql & ",codpobla=" & DBSet(vPr.CPostal, "T")
        sql = sql & ",pobprove=" & DBSet(vPr.Poblacion, "T")
        sql = sql & ",proprove=" & DBSet(vPr.Provincia, "T")
        sql = sql & ",nifprove=" & DBSet(vPr.NIF, "T")
        sql = sql & ",telprove=" & DBSet(vPr.TfnoAdmon, "T", "S")
        sql = sql & " WHERE"
        sql = sql & " numalbar = " & DBSet(Text1(0).Text, "T") & " AND fechaalb = " & DBSet(Text1(1).Text, "F") & " AND codprove = " & Text1(4).Text
        Conn.Execute sql
        Set vPr = Nothing
        If CadenaLineas <> "" Then
                'meto las lineas con el nuevo proveedor
                Conn.Execute CadenaLineas
                
                'UPDATEO las tablas de smoval
                sql = "UPDATE smoval SET codigope = " & NuevoProve
                sql = sql & " WHERE detamovi='ALC' AND "
                sql = sql & " document = " & DBSet(Text1(0).Text, "T") & " AND fechamov = " & DBSet(Text1(1).Text, "F") & " AND codigope = " & Text1(4).Text
                Conn.Execute sql
                
                
        End If
        
        
        HacerUpdatesCodProve = True
        Exit Function
EHacerUpdatesCodProve:
    MuestraError Err.Number, Err.Description
End Function


