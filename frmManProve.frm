VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmManProve 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Proveedores"
   ClientHeight    =   6165
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10725
   Icon            =   "frmManProve.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6165
   ScaleWidth      =   10725
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame4 
      Height          =   670
      Left            =   135
      TabIndex        =   40
      Top             =   540
      Width           =   10455
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   0
         Left            =   1320
         MaxLength       =   8
         TabIndex        =   0
         Tag             =   "Código Proveedor|N|N|0|999999|proveedor|codprove|000000|S|"
         Text            =   "Text1"
         Top             =   220
         Width           =   975
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   1
         Left            =   3345
         MaxLength       =   40
         TabIndex        =   1
         Tag             =   "Nombre Proveedor|T|N|||proveedor|nomprove||N|"
         Text            =   "Text1"
         Top             =   220
         Width           =   4245
      End
      Begin VB.CheckBox chkProveV 
         Caption         =   "Proveedor de Varios"
         Enabled         =   0   'False
         Height          =   195
         Index           =   0
         Left            =   8400
         TabIndex        =   2
         Tag             =   "Proveedor Varios|N|N|||proveedor|provario||N|"
         Top             =   220
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "Cod. Proveedor"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   42
         Top             =   220
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Nombre"
         Height          =   255
         Index           =   1
         Left            =   2640
         TabIndex        =   41
         Top             =   220
         Width           =   615
      End
   End
   Begin VB.Frame Frame1 
      Height          =   555
      Index           =   0
      Left            =   180
      TabIndex        =   35
      Top             =   5490
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
         TabIndex        =   36
         Top             =   180
         Width           =   2655
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   9420
      TabIndex        =   34
      Top             =   5610
      Width           =   1035
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   8130
      TabIndex        =   33
      Top             =   5610
      Width           =   1035
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   330
      Left            =   4185
      Top             =   5625
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
      TabIndex        =   38
      Top             =   0
      Width           =   10725
      _ExtentX        =   18918
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
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Buscar Tarjeta"
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
         TabIndex        =   39
         Top             =   120
         Width           =   1215
      End
   End
   Begin VB.CommandButton cmdRegresar 
      Caption         =   "&Regresar"
      Height          =   375
      Left            =   9405
      TabIndex        =   37
      Top             =   5625
      Visible         =   0   'False
      Width           =   1035
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4215
      Left            =   135
      TabIndex        =   43
      Top             =   1260
      Width           =   10455
      _ExtentX        =   18441
      _ExtentY        =   7435
      _Version        =   393216
      Style           =   1
      Tab             =   1
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "Datos básicos"
      TabPicture(0)   =   "frmManProve.frx":000C
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Label1(21)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "imgFec(1)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "imgFec(0)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "imgBuscar(0)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label1(19)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label1(13)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label1(14)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "imgBuscar(2)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label1(12)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label1(20)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "imgBuscar(1)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Label1(10)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Label1(11)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Label1(9)"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Label1(8)"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Label1(7)"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Label1(2)"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Label1(5)"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "Label1(3)"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "Label1(4)"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "Label1(6)"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "Text2(12)"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "Combo1(0)"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "Text1(5)"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "Text1(11)"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "Text2(14)"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "Text2(13)"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "Text1(12)"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "Text1(13)"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "Text1(18)"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "Text1(17)"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "Text1(16)"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "Text1(15)"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "Text1(14)"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).Control(34)=   "Combo1(1)"
      Tab(0).Control(34).Enabled=   0   'False
      Tab(0).Control(35)=   "Text1(10)"
      Tab(0).Control(35).Enabled=   0   'False
      Tab(0).Control(36)=   "Text1(9)"
      Tab(0).Control(36).Enabled=   0   'False
      Tab(0).Control(37)=   "Text1(8)"
      Tab(0).Control(37).Enabled=   0   'False
      Tab(0).Control(38)=   "Text1(7)"
      Tab(0).Control(38).Enabled=   0   'False
      Tab(0).Control(39)=   "Text1(2)"
      Tab(0).Control(39).Enabled=   0   'False
      Tab(0).Control(40)=   "Text1(3)"
      Tab(0).Control(40).Enabled=   0   'False
      Tab(0).Control(41)=   "Text1(4)"
      Tab(0).Control(41).Enabled=   0   'False
      Tab(0).Control(42)=   "Text1(6)"
      Tab(0).Control(42).Enabled=   0   'False
      Tab(0).Control(43)=   "Text1(29)"
      Tab(0).Control(43).Enabled=   0   'False
      Tab(0).ControlCount=   44
      TabCaption(1)   =   "Datos Contacto"
      TabPicture(1)   =   "frmManProve.frx":0028
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "imgWeb"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label2(11)"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label2(10)"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "imgZoom(0)"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Frame2(13)"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Text1(28)"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "Text1(27)"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "Frame1(2)"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).ControlCount=   8
      TabCaption(2)   =   "Documentos"
      TabPicture(2)   =   "frmManProve.frx":0044
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Text3(0)"
      Tab(2).Control(1)=   "Toolbar2"
      Tab(2).Control(2)=   "lw1"
      Tab(2).Control(3)=   "Toolbar3"
      Tab(2).Control(4)=   "Label16"
      Tab(2).Control(5)=   "Label17"
      Tab(2).Control(6)=   "imgFec1(0)"
      Tab(2).ControlCount=   7
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   29
         Left            =   -73320
         MaxLength       =   4
         TabIndex        =   9
         Tag             =   "IBAN|T|S|||proveedor|iban|||"
         Text            =   "Text1"
         Top             =   2445
         Width           =   615
      End
      Begin VB.TextBox Text3 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   0
         Left            =   -65850
         TabIndex        =   75
         Text            =   "Text4"
         Top             =   1080
         Width           =   1155
      End
      Begin VB.Frame Frame1 
         Caption         =   "Administración"
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
         Height          =   2080
         Index           =   2
         Left            =   180
         TabIndex        =   52
         Top             =   450
         Width           =   4935
         Begin VB.TextBox Text1 
            Height          =   315
            Index           =   19
            Left            =   120
            MaxLength       =   40
            TabIndex        =   23
            Tag             =   "Persona de Contacto Administración|T|S|||proveedor|perprov1|||"
            Text            =   "Text1"
            Top             =   600
            Width           =   4440
         End
         Begin VB.TextBox Text1 
            Height          =   315
            Index           =   20
            Left            =   135
            MaxLength       =   40
            TabIndex        =   24
            Tag             =   "eMail Administración|T|S|||proveedor|maiprov1|||"
            Text            =   "Text1"
            Top             =   1215
            Width           =   4440
         End
         Begin VB.TextBox Text1 
            Height          =   315
            Index           =   21
            Left            =   840
            MaxLength       =   15
            TabIndex        =   25
            Tag             =   "Telefono Administración|T|S|||proveedor|telprov1|||"
            Text            =   "Text1"
            Top             =   1640
            Width           =   1560
         End
         Begin VB.TextBox Text1 
            Height          =   315
            Index           =   22
            Left            =   3000
            MaxLength       =   15
            TabIndex        =   26
            Tag             =   "Fax Administración|T|S|||proveedor|faxprov1|||"
            Text            =   "Text1"
            Top             =   1640
            Width           =   1560
         End
         Begin VB.Label Label2 
            Caption         =   "Persona de Contacto"
            Height          =   240
            Index           =   13
            Left            =   120
            TabIndex        =   53
            Top             =   360
            Width           =   3495
         End
         Begin VB.Label Label2 
            Caption         =   "E-mail"
            Height          =   240
            Index           =   12
            Left            =   120
            TabIndex        =   54
            Top             =   960
            Width           =   435
         End
         Begin VB.Label Label2 
            Caption         =   "Teléfono"
            Height          =   240
            Index           =   1
            Left            =   120
            TabIndex        =   55
            Top             =   1640
            Width           =   735
         End
         Begin VB.Label Label2 
            Caption         =   "Fax"
            Height          =   240
            Index           =   0
            Left            =   2640
            TabIndex        =   56
            Top             =   1640
            Width           =   255
         End
         Begin VB.Image ImgMail 
            Height          =   240
            Index           =   0
            Left            =   600
            Tag             =   "-1"
            ToolTipText     =   "Enviar e-mail"
            Top             =   945
            Width           =   240
         End
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   6
         Left            =   -73320
         MaxLength       =   30
         TabIndex        =   7
         Tag             =   "Provincia|T|N|||proveedor|proprove|||"
         Text            =   "Text1"
         Top             =   1620
         Width           =   3270
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   4
         Left            =   -73320
         MaxLength       =   6
         TabIndex        =   5
         Tag             =   "CPostal|T|N|||proveedor|codpobla||N|"
         Text            =   "Text1"
         Top             =   1245
         Width           =   735
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   3
         Left            =   -73320
         MaxLength       =   35
         TabIndex        =   4
         Tag             =   "Domicilio|T|S|||proveedor|domprove||N|"
         Text            =   "Text1"
         Top             =   885
         Width           =   4230
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   2
         Left            =   -73320
         MaxLength       =   30
         TabIndex        =   3
         Tag             =   "Nombre Comercial|T|N|||proveedor|nomcomer||N|"
         Text            =   "Text1"
         Top             =   510
         Width           =   4245
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   7
         Left            =   -73320
         MaxLength       =   15
         TabIndex        =   8
         Tag             =   "N.I.F.|T|N|||proveedor|nifprove|||"
         Text            =   "Text1"
         Top             =   1995
         Width           =   2070
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   27
         Left            =   240
         MaxLength       =   40
         TabIndex        =   32
         Tag             =   "Web|T|S|||proveedor|wwwprove|||"
         Text            =   "Text1"
         Top             =   3735
         Width           =   6000
      End
      Begin VB.TextBox Text1 
         Height          =   580
         Index           =   28
         Left            =   240
         MaxLength       =   200
         TabIndex        =   31
         Tag             =   "Observaciones|T|S|||proveedor|observac|||"
         Text            =   "Text2 "
         Top             =   2850
         Width           =   9975
      End
      Begin VB.Frame Frame2 
         Caption         =   "Compras"
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
         Height          =   2080
         Index           =   13
         Left            =   5280
         TabIndex        =   47
         Top             =   450
         Width           =   4935
         Begin VB.TextBox Text1 
            Height          =   315
            Index           =   26
            Left            =   3000
            MaxLength       =   15
            TabIndex        =   30
            Tag             =   "Fax Compras|T|S|||proveedor|faxprov2|||"
            Text            =   "Text1"
            Top             =   1640
            Width           =   1560
         End
         Begin VB.TextBox Text1 
            Height          =   315
            Index           =   25
            Left            =   840
            MaxLength       =   15
            TabIndex        =   29
            Tag             =   "Teléfono Compras|T|S|||proveedor|telprov2|||"
            Text            =   "Text1"
            Top             =   1640
            Width           =   1560
         End
         Begin VB.TextBox Text1 
            Height          =   315
            Index           =   24
            Left            =   120
            MaxLength       =   40
            TabIndex        =   28
            Tag             =   "eMail Compras|T|S|||proveedor|maiprov2|||"
            Text            =   "Text1"
            Top             =   1200
            Width           =   4440
         End
         Begin VB.TextBox Text1 
            Height          =   315
            Index           =   23
            Left            =   120
            MaxLength       =   40
            TabIndex        =   27
            Tag             =   "Persona de Contacto Compras|T|S|||proveedor|perprov2|||"
            Text            =   "Text1"
            Top             =   600
            Width           =   4440
         End
         Begin VB.Label Label2 
            Caption         =   "Fax"
            Height          =   240
            Index           =   9
            Left            =   2640
            TabIndex        =   51
            Top             =   1640
            Width           =   255
         End
         Begin VB.Label Label2 
            Caption         =   "Teléfono"
            Height          =   240
            Index           =   8
            Left            =   120
            TabIndex        =   50
            Top             =   1640
            Width           =   735
         End
         Begin VB.Label Label2 
            Caption         =   "E-mail"
            Height          =   240
            Index           =   7
            Left            =   120
            TabIndex        =   49
            Top             =   960
            Width           =   480
         End
         Begin VB.Label Label2 
            Caption         =   "Persona de Contacto"
            Height          =   240
            Index           =   6
            Left            =   120
            TabIndex        =   48
            Top             =   360
            Width           =   3495
         End
         Begin VB.Image ImgMail 
            Height          =   240
            Index           =   1
            Left            =   600
            Tag             =   "-1"
            ToolTipText     =   "Enviar e-mail"
            Top             =   945
            Width           =   240
         End
      End
      Begin VB.TextBox Text1 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dd/MM/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   3
         EndProperty
         Height          =   315
         Index           =   8
         Left            =   -66480
         MaxLength       =   10
         TabIndex        =   18
         Tag             =   "Fecha de Alta|F|N|||proveedor|fecprove|dd/mm/yyyy||"
         Text            =   "Text1"
         Top             =   870
         Width           =   1575
      End
      Begin VB.TextBox Text1 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dd/MM/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   3
         EndProperty
         Height          =   315
         Index           =   9
         Left            =   -66480
         MaxLength       =   10
         TabIndex        =   19
         Tag             =   "Fecha última compra|F|S|||proveedor|fechamov|dd/mm/yyyy||"
         Text            =   "Text1"
         Top             =   1245
         Width           =   1575
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   1
         EndProperty
         Height          =   315
         Index           =   10
         Left            =   -66480
         MaxLength       =   5
         TabIndex        =   21
         Tag             =   "Dto. Pronto Pago|N|S|0|99.90|proveedor|dtoppago|#0.00||"
         Text            =   "Text1"
         Top             =   1980
         Width           =   735
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Index           =   1
         Left            =   -66480
         Style           =   2  'Dropdown List
         TabIndex        =   20
         Tag             =   "Tipo Descuento|N|N|||proveedor|tipodtos||N|"
         Top             =   1605
         Width           =   1575
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   14
         Left            =   -73320
         MaxLength       =   4
         TabIndex        =   16
         Tag             =   "Banco Propio|N|N|0|9999|proveedor|codbanpr|0000||"
         Text            =   "Text1"
         Top             =   3660
         Width           =   615
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   15
         Left            =   -72660
         MaxLength       =   4
         TabIndex        =   10
         Tag             =   "Banco|N|S|0|9999|proveedor|codbanco|0000||"
         Text            =   "Text1"
         Top             =   2445
         Width           =   615
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   16
         Left            =   -72000
         MaxLength       =   4
         TabIndex        =   11
         Tag             =   "Sucursal|N|S|0|9999|proveedor|codsucur|0000||"
         Text            =   "Text1"
         Top             =   2445
         Width           =   615
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   17
         Left            =   -71325
         MaxLength       =   2
         TabIndex        =   12
         Tag             =   "Digito Control|T|S|||proveedor|digcontr|00||"
         Text            =   "Text1"
         Top             =   2445
         Width           =   495
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   18
         Left            =   -70785
         MaxLength       =   10
         TabIndex        =   13
         Tag             =   "Cuenta Bancaria|T|S|||proveedor|cuentaba|0000000000||"
         Text            =   "Text1"
         Top             =   2445
         Width           =   1575
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   13
         Left            =   -73320
         MaxLength       =   3
         TabIndex        =   15
         Tag             =   "Forma Pago|N|N|0|999|proveedor|codforpa|000|N|"
         Text            =   "Text1"
         Top             =   3270
         Width           =   615
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   12
         Left            =   -73320
         MaxLength       =   10
         TabIndex        =   14
         Tag             =   "Cuenta Contable|T|S|||proveedor|codmacta|||"
         Text            =   "Text1"
         Top             =   2880
         Width           =   1335
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   315
         Index           =   13
         Left            =   -72690
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   46
         Text            =   "Text2"
         Top             =   3270
         Width           =   3135
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   315
         Index           =   14
         Left            =   -72690
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   45
         Text            =   "Text2"
         Top             =   3660
         Width           =   3135
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   1
         EndProperty
         Height          =   315
         Index           =   11
         Left            =   -66480
         MaxLength       =   5
         TabIndex        =   22
         Tag             =   "Dto. General|N|S|0|99.90|proveedor|dtognral|#0.00||"
         Text            =   "Text1"
         Top             =   2355
         Width           =   735
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   5
         Left            =   -71670
         MaxLength       =   30
         TabIndex        =   6
         Tag             =   "Población|T|N|||proveedor|pobprove||N|"
         Text            =   "Text1"
         Top             =   1245
         Width           =   2550
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Index           =   0
         Left            =   -66450
         TabIndex        =   17
         Tag             =   "Tipo de Proveedor|N|N|||proveedor|tipprove||N|"
         Text            =   "Combo1"
         Top             =   495
         Width           =   1575
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   315
         Index           =   12
         Left            =   -71970
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   44
         Text            =   "Text2"
         Top             =   2880
         Width           =   3735
      End
      Begin MSComctlLib.Toolbar Toolbar2 
         Height          =   1050
         Left            =   -74880
         TabIndex        =   76
         Top             =   540
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   1852
         ButtonWidth     =   609
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         Appearance      =   1
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   6
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Object.Visible         =   0   'False
               Object.ToolTipText     =   "Pedidos"
               Object.Tag             =   "0"
               Value           =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Object.Visible         =   0   'False
               Style           =   3
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Albaranes Compra"
               Object.Tag             =   "1"
               Style           =   2
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Facturas"
               Object.Tag             =   "2"
               Style           =   2
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ListView lw1 
         Height          =   3585
         Left            =   -74130
         TabIndex        =   77
         Top             =   510
         Width           =   7335
         _ExtentX        =   12938
         _ExtentY        =   6324
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
         Left            =   -74880
         TabIndex        =   78
         Top             =   540
         Visible         =   0   'False
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   3016
         ButtonWidth     =   609
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         Appearance      =   1
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   5
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Pedidos"
               Object.Tag             =   "0"
               Style           =   2
               Value           =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Albaranes Compra"
               Object.Tag             =   "1"
               Style           =   2
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Facturas"
               Style           =   2
            EndProperty
         EndProperty
      End
      Begin VB.Label Label16 
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
         Left            =   -66690
         TabIndex        =   80
         Top             =   600
         Width           =   1965
      End
      Begin VB.Label Label17 
         Caption         =   "Desde"
         Height          =   255
         Left            =   -66690
         TabIndex        =   79
         Top             =   1110
         Width           =   555
      End
      Begin VB.Image imgFec1 
         Height          =   240
         Index           =   0
         Left            =   -66120
         Picture         =   "frmManProve.frx":0060
         ToolTipText     =   "Buscar fecha"
         Top             =   1080
         Width           =   240
      End
      Begin VB.Image imgZoom 
         Height          =   240
         Index           =   0
         Left            =   1350
         ToolTipText     =   "Zoom descripción"
         Top             =   2565
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Provincia"
         Height          =   240
         Index           =   6
         Left            =   -74745
         TabIndex        =   57
         Top             =   1620
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Cod. postal"
         Height          =   255
         Index           =   4
         Left            =   -74745
         TabIndex        =   58
         Top             =   1245
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Domicilio"
         Height          =   255
         Index           =   3
         Left            =   -74745
         TabIndex        =   59
         Top             =   885
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Población"
         Height          =   255
         Index           =   5
         Left            =   -72465
         TabIndex        =   60
         Top             =   1245
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Nombre Comercial"
         Height          =   255
         Index           =   2
         Left            =   -74745
         TabIndex        =   74
         Top             =   510
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "N.I.F."
         Height          =   255
         Index           =   7
         Left            =   -74745
         TabIndex        =   73
         Top             =   1995
         Width           =   375
      End
      Begin VB.Label Label2 
         Caption         =   "Web"
         Height          =   240
         Index           =   10
         Left            =   240
         TabIndex        =   72
         Top             =   3520
         Width           =   375
      End
      Begin VB.Label Label2 
         Caption         =   "Observaciones"
         Height          =   240
         Index           =   11
         Left            =   240
         TabIndex        =   71
         Top             =   2630
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha de Alta"
         Height          =   195
         Index           =   8
         Left            =   -68145
         TabIndex        =   70
         Top             =   870
         Width           =   1080
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha Ult. Compra"
         Height          =   195
         Index           =   9
         Left            =   -68145
         TabIndex        =   69
         Top             =   1245
         Width           =   1320
      End
      Begin VB.Label Label1 
         Caption         =   "Cta Contable"
         Height          =   195
         Index           =   11
         Left            =   -74745
         TabIndex        =   68
         Top             =   2880
         Width           =   930
      End
      Begin VB.Label Label1 
         Caption         =   "Forma de Pago"
         Height          =   255
         Index           =   10
         Left            =   -74745
         TabIndex        =   67
         Top             =   3270
         Width           =   1095
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   1
         Left            =   -73605
         ToolTipText     =   "Buscar forma de pago"
         Top             =   3300
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Tipo Descuento"
         Height          =   255
         Index           =   20
         Left            =   -68145
         TabIndex        =   66
         Top             =   1605
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Dto. Pronto Pago"
         Height          =   195
         Index           =   12
         Left            =   -68145
         TabIndex        =   65
         Top             =   1980
         Width           =   1320
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   2
         Left            =   -73620
         ToolTipText     =   "Buscar banco propio"
         Top             =   3705
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Banco Propio"
         Height          =   195
         Index           =   14
         Left            =   -74745
         TabIndex        =   64
         Top             =   3660
         Width           =   1080
      End
      Begin VB.Label Label1 
         Caption         =   "Dto. General"
         Height          =   195
         Index           =   13
         Left            =   -68145
         TabIndex        =   63
         Top             =   2355
         Width           =   960
      End
      Begin VB.Label Label1 
         Caption         =   "Tipo Proveedor"
         Height          =   255
         Index           =   19
         Left            =   -68145
         TabIndex        =   62
         Top             =   495
         Width           =   1110
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   0
         Left            =   -73620
         Tag             =   "-1"
         ToolTipText     =   "Buscar cuenta contable"
         Top             =   2925
         Width           =   240
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   0
         Left            =   -66765
         Picture         =   "frmManProve.frx":00EB
         ToolTipText     =   "Buscar fecha"
         Top             =   840
         Width           =   240
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   1
         Left            =   -66765
         Picture         =   "frmManProve.frx":0176
         ToolTipText     =   "Buscar fecha"
         Top             =   1245
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "IBAN Proveedor"
         Height          =   195
         Index           =   21
         Left            =   -74745
         TabIndex        =   61
         Top             =   2505
         Width           =   1320
      End
      Begin VB.Image imgWeb 
         Height          =   255
         Left            =   720
         Picture         =   "frmManProve.frx":0201
         Stretch         =   -1  'True
         Tag             =   "-1"
         ToolTipText     =   "Abrir web"
         Top             =   3475
         Width           =   255
      End
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
      Begin VB.Menu mnBarra3 
         Caption         =   "-"
      End
      Begin VB.Menu mnSalir 
         Caption         =   "&Salir"
         Shortcut        =   ^S
      End
   End
End
Attribute VB_Name = "frmManProve"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' +-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+
' +-+- Autor: MANOLO                   -+-+
' +-+- Menú: PROVEEDORES               -+-+
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
Private WithEvents frmC As frmCal 'calendario fecha
Attribute frmC.VB_VarHelpID = -1
Private WithEvents frmC1 As frmCal 'calendario fecha
Attribute frmC1.VB_VarHelpID = -1
Private WithEvents frmZ As frmZoom  'Zoom para campos Text
Attribute frmZ.VB_VarHelpID = -1

Private WithEvents frmBan As frmManBanco 'Banco Propio
Attribute frmBan.VB_VarHelpID = -1
Private WithEvents frmFPa As frmManFpago   'Formas de Pago
Attribute frmFPa.VB_VarHelpID = -1

Private WithEvents frmCtas As frmCtasConta 'cuentas contables
Attribute frmCtas.VB_VarHelpID = -1
Private WithEvents frmFac As frmComHcoFacturas ' hco de facturas de proveedores
Attribute frmFac.VB_VarHelpID = -1
Private WithEvents frmAlb As frmComEntAlbaranes ' albaranes de compra
Attribute frmAlb.VB_VarHelpID = -1
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
Private BuscaChekc As String

'Private VieneDeBuscar As Boolean
'Per a quan torna 2 poblacions en el mateix codi Postal. Si ve de pulsar prismatic
'de búsqueda posar el valor de població seleccionada i no tornar a recuperar de la Base de Datos

Dim btnPrimero As Byte 'Variable que indica el nº del Botó PrimerRegistro en la Toolbar1
'Dim CadAncho() As Boolean  'array, per a quan cridem al form de llínies
Dim indice As Byte 'Index del text1 on es poses els datos retornats des d'atres Formularis de Mtos
Dim CadB As String

Dim NombreAnt As String
Dim DomAnt As String
Dim EMaiAnt As String
Dim WebAnt As String

Dim cPostalAnt As String
Dim PobAnt As String
Dim ProAnt As String
Dim NifAnt As String
Dim TelAnt As String
Dim ForpaAnt As String
Dim IbanAnt As String

Dim BancoAnt  As String
Dim SucurAnt As String
Dim DigitoAnt As String
Dim CuentaAnt As String


Private Sub cmbAux_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub chkProveV_KeyPress(Index As Integer, KeyAscii As Integer)
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
                    Text2(12).Text = PonerNombreCuenta(Text1(12), Modo, Text1(0).Text)
        
                    Data1.RecordSource = "Select * from " & NombreTabla & Ordenacion
                    PosicionarData
                End If
            Else
                ModoLineas = 0
            End If
            
        Case 4  'MODIFICAR
            If DatosOk Then
                Modificar
'                If ModificaDesdeFormulario2(Me, 1) Then
'                    TerminaBloquear
'                    PosicionarData
'                End If
            Else
                ModoLineas = 0
            End If
    End Select
    Screen.MousePointer = vbDefault
 
Error1:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub

' *** si n'hi han combos a la capçalera ***
Private Sub Combo1_GotFocus(Index As Integer)
    If Modo = 1 Then Combo1(Index).BackColor = vbYellow
End Sub

Private Sub Combo1_LostFocus(Index As Integer)
    If Combo1(Index).BackColor = vbYellow Then Combo1(Index).BackColor = vbWhite
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
    If Modo = 4 Then TerminaBloquear
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
        .Buttons(11).Image = 21   'Borrar
        'el 10 i el 11 son separadors
        .Buttons(12).Image = 10  'Imprimir
        .Buttons(13).Image = 11  'Eixir
        'el 13 i el 14 son separadors
        .Buttons(btnPrimero).Image = 6  'Primer
        .Buttons(btnPrimero + 1).Image = 7 'Anterior
        .Buttons(btnPrimero + 2).Image = 8 'Següent
        .Buttons(btnPrimero + 3).Image = 9 'Últim
    End With
    
    
    'cargar IMAGES de busqueda
    For i = 0 To Me.imgBuscar.Count - 1
        Me.imgBuscar(i).Picture = frmPpal.imgListImages16.ListImages(1).Picture
    Next i
    
    'carga IMAGES de mail
    For i = 0 To Me.imgMail.Count - 1
        Me.imgMail(i).Picture = frmPpal.imgListImages16.ListImages(2).Picture
    Next i
    
    'IMAGES para zoom
    For i = 0 To Me.imgZoom.Count - 1
        Me.imgZoom(i).Picture = frmPpal.imgListImages16.ListImages(3).Picture
    Next i
   
   'La nevegacion para entradas, facturas....
    ImagenesNavegacion
   'Ponemos los datos del listview
    imgFec(0).Tag = "01/01/" & Year(Now) 'vParam.FecIniCam
    CargaColumnas 1 '0
    
    
    
    ' *** si n'hi han tabs, per a que per defecte sempre es pose al 1r***
    Me.SSTab1.Tab = 0
    ' *******************************************************************
    
    LimpiarCampos   'Neteja els camps TextBox
    
    '*** canviar el nom de la taula i l'ordenació de la capçalera ***
    NombreTabla = "proveedor"
    Ordenacion = " ORDER BY codprove"
    
    'Mirem com està guardat el valor del check
    chkVistaPrevia(0).Value = CheckValueLeer(Name)
    
    Data1.ConnectionString = Conn
    '***** cambiar el nombre de la PK de la cabecera *************
    Data1.RecordSource = "Select * from " & NombreTabla & " where codprove=-1"
    Data1.Refresh
       
    ModoLineas = 0
       
    CargaCombo
    
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
    
    ' limpiamos los combos
    Combo1(0).ListIndex = -1
    Combo1(1).ListIndex = -1
    
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
Dim i As Integer, Numreg As Byte
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
    For i = 0 To imgFec.Count - 1
        BloquearImgFec Me, i, Modo
    Next i
    ' ********************************************************
    
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
    
    chkVistaPrevia(0).Enabled = (Modo <= 2)
    
    PonerLongCampos

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
    
End Sub

Private Sub Desplazamiento(Index As Integer)
'Botons de Desplaçament; per a desplaçar-se pels registres de control Data
    If Data1.Recordset.EOF Then Exit Sub
    DesplazamientoData Data1, Index
    PonerCampos
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

Private Sub frmBan_DatoSeleccionado(CadenaSeleccion As String)
'Mantenimiento Bancos Propios
    Text1(14).Text = RecuperaValor(CadenaSeleccion, 1) 'codbanpr
    FormateaCampo Text1(14)
    Text2(14).Text = RecuperaValor(CadenaSeleccion, 2) 'nombanpr

End Sub

Private Sub frmC1_Selec(vFecha As Date)
    ' *** repasar si el camp es txtAux o Text1 ***
    Text3(0).Text = Format(vFecha, "dd/mm/yyyy")  '<===
    ' ********
End Sub

Private Sub frmZ_Actualizar(vCampo As String)
     Text1(indice).Text = vCampo
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

    frmC.Left = esq + imgFec(Index).Parent.Left + 30
    frmC.Top = dalt + imgFec(Index).Parent.Top + imgFec(Index).Height + menu - 40

    imgFec(0).Tag = Index + 8 '<===
    ' *** repasar si el camp es txtAux o Text1 ***
    If Text1(Index + 8).Text <> "" Then frmC.NovaData = Text1(Index + 8).Text
    ' ********************************************

    frmC.Show vbModal
    Set frmC = Nothing
    ' *** repasar si el camp es txtAux o Text1 ***
    PonerFoco Text1(CByte(imgFec(0).Tag)) '<===
    ' ********************************************
End Sub

Private Sub imgFec1_Click(Index As Integer)
    Dim esq As Long
    Dim dalt As Long
    Dim menu As Long
    Dim obj As Object

    Set frmC1 = New frmCal
    
    esq = imgFec1(Index).Left
    dalt = imgFec1(Index).Top
        
    Set obj = imgFec1(Index).Container
      
      While imgFec1(Index).Parent.Name <> obj.Name
            esq = esq + obj.Left
            dalt = dalt + obj.Top
            Set obj = obj.Container
      Wend
    
    menu = Me.Height - Me.ScaleHeight 'ací tinc el heigth del menú i de la toolbar

    frmC1.Left = esq + imgFec1(Index).Parent.Left + 30
    frmC1.Top = dalt + imgFec1(Index).Parent.Top + imgFec1(Index).Height + menu - 40

    imgFec1(0).Tag = Index '<===
    ' *** repasar si el camp es txtAux o Text1 ***
    If Text3(0).Text <> "" Then frmC1.NovaData = Text3(0).Text
    ' ********************************************

    frmC1.Show vbModal
    Set frmC1 = Nothing
    ' *** repasar si el camp es txtAux o Text1 ***
    PonerFoco Text3(0) '<===
    ' ********************************************
    
    If Text3(0).Text <> "" Then
         imgFec1(0).Tag = Text3(0).Text
         CargaDatosLW
    End If
    
End Sub



Private Sub frmC_Selec(vFecha As Date)
    ' *** repasar si el camp es txtAux o Text1 ***
    Text1(CByte(imgFec(0).Tag)).Text = Format(vFecha, "dd/mm/yyyy")  '<===
    ' ********************************************
End Sub
' *****************************************************


Private Sub imgMail_Click(Index As Integer)
    Select Case Index
        Case 0
            If Text1(20).Text <> "" Then
                LanzaMailGnral Text1(20).Text
            End If
        Case 1
            If Text1(24).Text <> "" Then
                LanzaMailGnral Text1(24).Text
            End If
    End Select
End Sub

Private Sub imgZoom_Click(Index As Integer)
    
    Set frmZ = New frmZoom

    If Index = 0 Then
        indice = 28
        frmZ.pTitulo = "Observaciones del Proveedor"
        frmZ.pValor = Text1(indice).Text
        frmZ.pModo = Modo
    
        frmZ.Show vbModal
        Set frmZ = Nothing
            
        PonerFoco Text1(indice)
    End If
End Sub

Private Sub mnBuscar_Click()
    BotonBuscar
End Sub

Private Sub mnEliminar_Click()
    BotonEliminar
End Sub

Private Sub mnImprimir_Click()
    AbrirListado (15)
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
'            mnBuscarTarjeta_Click
        Case 12 'Imprimir
            mnImprimir_Click
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
        Combo1(0).ListIndex = -1 'quan busque, per defecte no seleccione cap tipo de proveedor
        Combo1(1).ListIndex = -1
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

    CadB = ObtenerBusqueda2(Me, 0)
    
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
    Dim Cad As String
        
    'Cridem al form
    ' **************** arreglar-ho per a vore lo que es desije ****************
    ' NOTA: el total d'amples de ParaGrid, ha de sumar 100
    Cad = ""
    Cad = Cad & ParaGrid(Text1(0), 15, "Cód.")
    Cad = Cad & ParaGrid(Text1(1), 60, "Nombre")
    Cad = Cad & ParaGrid(Text1(2), 25, "N.I.F.")
    If Cad <> "" Then
        Screen.MousePointer = vbHourglass
        Set frmB = New frmBuscaGrid
        frmB.vCampos = Cad
        frmB.vTabla = NombreTabla
        frmB.vSQL = CadB
        HaDevueltoDatos = False
        frmB.vDevuelve = "0|1|2|" '*** els camps que volen que torne ***
        frmB.vTitulo = "Proveedor" ' ***** repasa açò: títol de BuscaGrid *****
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
Dim Cad As String
Dim Aux As String
Dim i As Integer
Dim J As Integer

    If Data1.Recordset.EOF Then
        MsgBox "Ningún registro devuelto.", vbExclamation
        Exit Sub
    End If
    
    Cad = ""
    i = 0
    Do
        J = i + 1
        i = InStr(J, DatosADevolverBusqueda, "|")
        If i > 0 Then
            Aux = Mid(DatosADevolverBusqueda, J, i - J)
            J = Val(Aux)
            Cad = Cad & Text1(J).Text & "|"
        End If
    Loop Until i = 0
    RaiseEvent DatoSeleccionado(Cad)
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
    Text1(0).Text = SugerirCodigoSiguienteStr("proveedor", "codprove")
    FormateaCampo Text1(0)
      
    PosicionarCombo Combo1(0), 0
    PosicionarCombo Combo1(1), 0
       
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
    
    NombreAnt = Text1(1).Text
    DomAnt = Text1(3).Text
    cPostalAnt = Text1(4).Text
    PobAnt = Text1(5).Text
    ProAnt = Text1(6).Text
    NifAnt = Text1(7).Text
    TelAnt = Text1(21).Text
    EMaiAnt = Text1(20).Text
    WebAnt = Text1(27).Text
    
    IbanAnt = Text1(29).Text
    BancoAnt = Text1(15).Text
    SucurAnt = Text1(16).Text
    DigitoAnt = Text1(17).Text
    CuentaAnt = Text1(18).Text
    ForpaAnt = Text1(13).Text
    
    ' *** foco al 1r camp visible que NO siga clau primaria ***
    PonerFoco Text1(1)
End Sub

Private Sub BotonEliminar()
Dim Cad As String

    On Error GoTo EEliminar

    'Ciertas comprobaciones
    If Data1.Recordset.EOF Then Exit Sub

    ' *** repasar el nom de l'adodc, l'index del Field i el camp que te la PK ***
    ' ### [Monica] 26/09/2006 dejamos modificar y eliminar el codigo 0
    'El registre de codi 0 no es pot Modificar ni Eliminar
    'If EsCodigoCero(CStr(Data1.Recordset.Fields(0).Value), FormatoCampo(text1(0))) Then Exit Sub
    ' ***************************************************************************

    ' *************** canviar la pregunta ****************
    Cad = "¿Seguro que desea eliminar el Proveedor?"
    Cad = Cad & vbCrLf & "Código: " & Format(Data1.Recordset.Fields(0), FormatoCampo(Text1(0)))
    Cad = Cad & vbCrLf & "Nombre: " & Data1.Recordset.Fields(1)
    
    If MsgBox(Cad, vbQuestion + vbYesNo) = vbYes Then
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
    
    
    ' ************* configurar els camps de les descripcions de la capçalera *************
    Text2(14).Text = PonerNombreDeCod(Text1(14), "sbanco", "nombanco")
    Text2(13).Text = PonerNombreDeCod(Text1(13), "sforpa", "nomforpa", "codforpa", "N")
    If vParamAplic.NumeroConta <> 0 Then
        Text2(12).Text = PonerNombreCuenta(Text1(12), Modo)
    End If
    ' ********************************************************************************
    
    lblIndicador.Caption = "Datos navegacion"
    Me.Refresh
    DoEvents
    CargaDatosLW
    
    
    '-- Esto permanece para saber donde estamos
    lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
    
'    PonerModoOpcionesMenu (Modo)
'    PonerOpcionesMenu
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
        
    End Select
End Sub

Private Function DatosOk() As Boolean
Dim b As Boolean
'Dim Datos As String
Dim cta As String
Dim cadMen As String


    On Error GoTo EDatosOK

    DatosOk = False
    b = CompForm2(Me, 1)
    If Not b Then Exit Function
    
    ' *** canviar els arguments de la funcio, el mensage i repasar si n'hi ha codEmpre ***
    If (Modo = 3) Then 'insertar
        'comprobar si existe ya el cod. del campo clave primaria
        If ExisteCP(Text1(0)) Then b = False
    End If
    
    If b And (Modo = 3 Or Modo = 4) Then
        '[Monica]22/11/2013: añadida la comprobacion de que la cuenta contable sea correcta
        If Text1(15).Text = "" Or Text1(16).Text = "" Or Text1(17).Text = "" Or Text1(18).Text = "" Then
            '[Monica]20/11/2013: añadido el codigo de iban
            Text1(29).Text = ""
            Text1(15).Text = ""
            Text1(16).Text = ""
            Text1(17).Text = ""
            Text1(18).Text = ""
        Else
            cta = Format(Text1(15).Text, "0000") & Format(Text1(16).Text, "0000") & Format(Text1(17).Text, "00") & Format(Text1(18).Text, "0000000000")
            If Val(ComprobarCero(cta)) = 0 Then
                cadMen = "El proveedor no tiene asignada cuenta bancaria."
                MsgBox cadMen, vbExclamation
            End If
            If Not Comprueba_CC(cta) Then
                cadMen = "La cuenta bancaria del proveedor no es correcta. ¿ Desea continuar ?."
                If MsgBox(cadMen, vbQuestion + vbYesNo + vbDefaultButton1) = vbYes Then
                    b = True
                Else
                    PonerFoco Text1(15)
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
                If Me.Text1(29).Text <> "" Then BuscaChekc = Mid(Text1(29).Text, 1, 2)
                    
                If DevuelveIBAN2(BuscaChekc, cta, cta) Then
                    If Me.Text1(29).Text = "" Then
                        If MsgBox("Poner IBAN ?", vbQuestion + vbYesNo) = vbYes Then Me.Text1(29).Text = BuscaChekc & cta
                    Else
                        If Mid(Text1(29).Text, 3) <> cta Then
                            cta = "Calculado : " & BuscaChekc & cta
                            cta = "Introducido: " & Me.Text1(29).Text & vbCrLf & cta & vbCrLf
                            cta = "Error en codigo IBAN" & vbCrLf & cta & "Continuar?"
                            If MsgBox(cta, vbQuestion + vbYesNo) = vbNo Then
                                PonerFoco Text1(29)
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
Dim Cad As String, Indicador As String

    ' *** canviar-ho per tota la PK de la capçalera, no llevar els () ***
    Cad = "(codprove=" & Text1(0).Text & ")"
    
    ' *** gastar SituarData o SituarDataMULTI depenent de si la PK es simple o composta ***
    'If SituarDataMULTI(Data1, cad, Indicador) Then
    If SituarData(Data1, Cad, Indicador) Then
        If ModoLineas <> 1 Then PonerModo 2
        lblIndicador.Caption = Indicador
    Else
       LimpiarCampos
       PonerModo 0
    End If
End Sub

Private Function Eliminar() As Boolean
Dim vWhere As String

    On Error GoTo FinEliminar

    Conn.BeginTrans
    ' ***** canviar el nom de la PK de la capçalera, repasar codEmpre *******
    vWhere = " WHERE codprove=" & Data1.Recordset!CodProve
        
    'Eliminar la CAPÇALERA
    Conn.Execute "Delete from " & NombreTabla & vWhere
       
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
Dim cadMen As String
Dim Nuevo As Boolean

    If Not PerderFocoGnral(Text1(Index), Modo) Then Exit Sub
    
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
'    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub
    
    ' ***************** configurar els LostFocus dels camps de la capçalera *****************
    Select Case Index
        Case 0 'PROVEEDOR
            PonerFormatoEntero Text1(Index)

        Case 1, 2 'NOMBRE y NOMBRE COMERCIAL
            Text1(Index).Text = UCase(Text1(Index).Text)
        
        Case 7 'NIF
            Text1(Index).Text = UCase(Text1(Index).Text)
            ValidarNIF Text1(Index).Text
                
                
        Case 14 'BANCO PROPIO
            If PonerFormatoEntero(Text1(Index)) Then
                Text2(Index).Text = PonerNombreDeCod(Text1(Index), "sbanco", "nombanco")
                If Text2(Index).Text = "" Then
                    cadMen = "No existe el Banco Propio: " & Text1(Index).Text & vbCrLf
                    cadMen = cadMen & "¿Desea crearlo?" & vbCrLf
                    If MsgBox(cadMen, vbQuestion + vbYesNo) = vbYes Then
                        Set frmBan = New frmManBanco
                        frmBan.DatosADevolverBusqueda = "0|1|"
                        frmBan.NuevoCodigo = Text1(Index).Text
                        Text1(Index).Text = ""
                        TerminaBloquear
                        frmBan.Show vbModal
                        Set frmBan = Nothing
                        If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Data1, 1
                    Else
                        Text1(Index).Text = ""
                    End If
                    PonerFoco Text1(Index)
                End If
            Else
                Text2(Index).Text = ""
            End If
        
        Case 13 'FORMA DE PAGO
            If PonerFormatoEntero(Text1(Index)) Then
                Text2(Index).Text = PonerNombreDeCod(Text1(Index), "sforpa", "nomforpa")
                If Text2(Index).Text = "" Then
                    cadMen = "No existe la Forma de Pago: " & Text1(Index).Text & vbCrLf
                    cadMen = cadMen & "¿Desea crearla?" & vbCrLf
                    If MsgBox(cadMen, vbQuestion + vbYesNo) = vbYes Then
                        Set frmFPa = New frmManFpago
                        frmFPa.DatosADevolverBusqueda = "0|1|"
                        frmFPa.NuevoCodigo = Text1(Index).Text
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
                Text2(Index).Text = ""
            End If
            
        Case 10, 11 'Dto.Pronto Pago, Dto.General
            If Text1(Index).Text = "" Then Exit Sub
            PonerFormatoDecimal Text1(Index), 4
            
        Case 15, 16 'ENTIDAD Y SUCURSAL BANCARIA
            PonerFormatoEntero Text1(Index)
        
        Case 8, 9 'Fechas
            PonerFormatoFecha Text1(Index)
            
        Case 12 'cuenta contable
            If Text1(Index).Text = "" Then Exit Sub
            If Modo = 3 Then
                Text2(Index).Text = PonerNombreCuenta(Text1(Index), Modo, "") 'text1(0).Text)
            Else
                Text2(Index).Text = PonerNombreCuenta(Text1(Index), Modo, Text1(0).Text)
            End If
            
        Case 29 ' codigo de iban
            Text1(Index).Text = UCase(Text1(Index).Text)
            
    End Select
    
    '[Monica]: calculo del iban si no lo ponen
    If Index = 15 Or Index = 16 Or Index = 17 Or Index = 18 Then
        Dim cta As String
        Dim CC As String
        If Text1(15).Text <> "" And Text1(16).Text <> "" And Text1(17).Text <> "" And Text1(18).Text <> "" Then
            
            cta = Format(Text1(15).Text, "0000") & Format(Text1(16).Text, "0000") & Format(Text1(17).Text, "00") & Format(Text1(18).Text, "0000000000")
            If Len(cta) = 20 Then
    '        Text1(42).Text = Calculo_CC_IBAN(cta, Text1(42).Text)
    
                If Text1(29).Text = "" Then
                    'NO ha puesto IBAN
                    If DevuelveIBAN2("ES", cta, cta) Then Text1(29).Text = "ES" & cta
                Else
                    CC = CStr(Mid(Text1(29).Text, 1, 2))
                    If DevuelveIBAN2(CStr(CC), cta, cta) Then
                        If Mid(Text1(29).Text, 3) <> cta Then
                            
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
                Case 14: KEYBusqueda KeyAscii, 2 'banco propio
                Case 12: KEYBusqueda KeyAscii, 0 'cuenta contable
                Case 13: KEYBusqueda KeyAscii, 1 'forma pago
                Case 8: KEYFecha KeyAscii, 8 'fecha de alta
                Case 12: KEYFecha KeyAscii, 12 'fecha de baja
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
        Case 2 'banco propio
            Set frmBan = New frmManBanco
            frmBan.DatosADevolverBusqueda = "0|1|"
            frmBan.CodigoActual = Text1(14).Text
            frmBan.Show vbModal
            Set frmBan = Nothing
            PonerFoco Text1(14)
            
        Case 1 'formas de pago
            Set frmFPa = New frmManFpago
            frmFPa.DatosADevolverBusqueda = "0|1|"
            frmFPa.CodigoActual = Text1(13).Text
            frmFPa.Show vbModal
            Set frmFPa = Nothing
            PonerFoco Text1(13)
            
        Case 0 'Cuentas Contables (de contabilidad)
            If vParamAplic.NumeroConta = 0 Then Exit Sub
            
            indice = Index + 12
            Set frmCtas = New frmCtasConta
            frmCtas.NumDigit = 0
            frmCtas.DatosADevolverBusqueda = "0|1|"
            frmCtas.CodigoActual = Text1(indice).Text
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

Private Sub frmFpa_DatoSeleccionado(CadenaSeleccion As String)
'Mantenimiento F.Pago
    Text1(13).Text = RecuperaValor(CadenaSeleccion, 1) 'codforpa
    FormateaCampo Text1(13)
    Text2(13).Text = RecuperaValor(CadenaSeleccion, 2) 'nomforpa
End Sub


Private Sub imgWeb_Click()
    'Abrimos el explorador de windows con la pagina Web del proveedor
    If Modo = 0 Then Exit Sub
    Screen.MousePointer = vbHourglass
'    If LanzaHome("websoporte") Then espera 2
    If LanzaHomeGnral(Text1(10).Text) Then espera 2
    Screen.MousePointer = vbDefault
End Sub

Private Function ObtenerWhereCab(conW As Boolean) As String
Dim vWhere As String
    
    vWhere = ""
    If conW Then vWhere = " WHERE "
    ' *** canviar-ho per la clau primaria de la capçalera ***
    vWhere = vWhere & " codsocio=" & Val(Text1(0).Text)
    
    ObtenerWhereCab = vWhere
End Function

' ********* si n'hi han combos a la capçalera ************
Private Sub CargaCombo()
Dim Ini As Integer
Dim Fin As Integer
Dim i As Integer

    ' *** neteje els combos, els pose valor i seleccione el valor per defecte ***
    For i = 0 To Combo1.Count - 1
        Combo1(i).Clear
    Next i
    
    Combo1(0).AddItem "Nacional"
    Combo1(0).ItemData(Combo1(0).NewIndex) = 0
    Combo1(0).AddItem "Intracomunitario"
    Combo1(0).ItemData(Combo1(0).NewIndex) = 1
    Combo1(0).AddItem "Extranjero"
    Combo1(0).ItemData(Combo1(0).NewIndex) = 2
    
    Combo1(1).AddItem "Aditivo"
    Combo1(1).ItemData(Combo1(1).NewIndex) = 0
    Combo1(1).AddItem "Resto"
    Combo1(1).ItemData(Combo1(1).NewIndex) = 1
    
End Sub

'-----------------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------------
'
'
'       El listview tendra los datos de albaranes, facturas... que tenga el cliente
'       Con lo cual, a partir de un click tendremos que ser capaces de situarnos en
'       el formulario correspondiente
'
'-----------------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------------


Private Sub ImagenesNavegacion()
    With Me.Toolbar2
        .ImageList = frmPpal.imgListPpal
        .Buttons(1).Image = 1 ' 15
        .Buttons(3).Image = 10
        .Buttons(5).Image = 2 '22
    End With
    'tenemos un toolbar oculto para que el icono sea de 16
    With Me.Toolbar3
        .ImageList = frmPpal.imgListImages16
        .Buttons(1).Image = 1 '13
        .Buttons(3).Image = 1 '8
        .Buttons(5).Image = 1 '7
    End With
    
'    Set lw1.SmallIcons = frmPpal.imgListPpal
    Set lw1.SmallIcons = frmPpal.imgListImages16
    
End Sub


Private Sub Toolbar2_ButtonClick(ByVal Button As MSComctlLib.Button)
    
    If Button.Tag = "" Then Exit Sub
    Label16.Caption = ""
    'Levantamos todos los botones y dejamos pulsado el de ahora
    For NumRegElim = 1 To Toolbar2.Buttons.Count
        If Toolbar2.Buttons(NumRegElim).Tag <> "" Then
            If Toolbar2.Buttons(NumRegElim).Index <> Button.Index Then Toolbar2.Buttons(NumRegElim).Value = tbrUnpressed
        End If
    Next NumRegElim
    CargaColumnas CByte(Button.Tag)
    
    'Hacemos las acciones
    If Modo = 2 Then CargaDatosLW
End Sub



Private Sub CargaColumnas(OpcionList As Byte)
Dim Columnas As String
Dim Ancho As String
Dim Alinea As String
Dim Formato As String
Dim Ncol As Integer
Dim C As ColumnHeader

    Select Case OpcionList
    Case 0
        'Pedidos
        Label16.Caption = "Pedidos Compra"
        Columnas = "Pedido|Fecha|Forma Pago|"
        Ancho = "1200|1200|2300|"
        'vwColumnRight =1  left=0   center=2
        Alinea = "0|0|0|"
        'Formatos
        Formato = "|dd/mm/yyyy|0|"
        Ncol = 3
               
    Case 1
        'Albaranes
        Label16.Caption = "Albaranes Compra"
        Columnas = "Albarán|Fecha|Forma Pago|Pedido|Fec.Pedido|"
        Ancho = "1200|1200|2300|1000|1100|"
        'vwColumnRight =1  left=0   center=2
        Alinea = "0|0|0|0|0|"
        'Formatos
        Formato = "|dd/mm/yyyy|0|0|dd/mm/yyyy|"
        Ncol = 5
    
    Case 2
        'Facturas
        Label16.Caption = "Facturas"
        Columnas = "Tipo|Numero|Fecha|Importe|"
        Ancho = "1000|2000|1200|2700|"
        'vwColumnRight =1  left=0   center=2
        Alinea = "0|0|0|1|"
        'Formatos
        Formato = "|00000000|dd/mm/yyyy|" & FormatoImporte & "|"
        Ncol = 4
    
    End Select
    
    
'    'Fecha incio busquedas
'    Text1(46).Text = Format(imgFecha(3).Tag, "dd/mm/yyyy")
    'Guardo la opcion en el tag
    lw1.Tag = OpcionList & "|" & Ncol & "|"
    
    lw1.ColumnHeaders.Clear
    
    For NumRegElim = 1 To Ncol
         Set C = lw1.ColumnHeaders.Add()
         C.Text = RecuperaValor(Columnas, CInt(NumRegElim))
         C.Width = RecuperaValor(Ancho, CInt(NumRegElim))
         C.Alignment = Val(RecuperaValor(Alinea, CInt(NumRegElim)))
         C.Tag = RecuperaValor(Formato, CInt(NumRegElim))
    Next NumRegElim
End Sub


Private Sub CargaDatosLW()
Dim C As String
Dim bs As Byte
    bs = Screen.MousePointer
    C = Me.lblIndicador.Caption
    lblIndicador.Caption = "Leyendo " & Label16.Caption
    lblIndicador.Refresh
    CargaDatosLW2
    Me.lblIndicador.Caption = C
    Screen.MousePointer = bs
End Sub

Private Sub CargaDatosLW2()
Dim Cad As String
Dim Rs As ADODB.Recordset
Dim It As ListItem
Dim ElIcono As Integer
Dim GroupBy As String
Dim Orden As String


    On Error GoTo ECargaDatosLW
    
    If Modo <> 2 Then Exit Sub
    
    For NumRegElim = 1 To Toolbar2.Buttons.Count
        If Toolbar2.Buttons(NumRegElim).Value = tbrPressed Then
            ElIcono = Toolbar3.Buttons(NumRegElim).Image
            Exit For
        End If
    Next
    
    
    'Fecha incio busquedas
    Text3(0).Text = Format(imgFec1(0).Tag, "dd/mm/yyyy")
    
    
    Select Case CByte(RecuperaValor(lw1.Tag, 1))
    Case 0
'        'PEDIDOS
'        cad = "select h.numpedpr,h.fecpedpr,f.nomforpa from scappr h, sforpa f WHERE "
'        cad = cad & " h.codforpa = f.codforpa "
'        GroupBy = "1,2,3 "
'        BuscaChekc = "h.fecpedpr"
        
    Case 1
        'ALBARANES DE compra
        Cad = "select h.numalbar,h.fechaalb,f.nomforpa,h.numpedpr,h.fecpedpr from scaalp h, sforpa f where "
        Cad = Cad & " h.codforpa = f.codforpa "
        
        GroupBy = "1,2,3,4,5 "
        BuscaChekc = "h.fechaalb"
        Orden = "h.numalbar"
    
    Case 2
        'FACTURAS
        Cad = "select h.numfactu,h.fecfactu,h.fecrecep,h.totalfac from scafpc h WHERE 1=1"
        GroupBy = "1,2,3,4 "
        BuscaChekc = "h.fecfactu"
    End Select
    
    
    'La fecha
    
    'EL where del codclien
    Cad = Cad & " and h.codprove=" & Data1.Recordset!CodProve
    
    'La fecha
    If BuscaChekc <> "" Then Cad = Cad & " and " & BuscaChekc & " >='" & Format(imgFec1(0).Tag, FormatoFecha) & "'"
    
    
    'El group by
    If GroupBy <> "" Then Cad = Cad & " GROUP BY " & GroupBy
    
    'El ORDER BY
'    If CByte(RecuperaValor(lw1.Tag, 1)) = 1 Then BuscaChekc = Orden
    
    'BuscaChekc="" si es la opcion de precios especiales
    Cad = Cad & " ORDER BY " & BuscaChekc & " DESC"
    BuscaChekc = ""
    
    lw1.ListItems.Clear
    Set Rs = New ADODB.Recordset
    Rs.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not Rs.EOF
        Set It = lw1.ListItems.Add()
        If lw1.ColumnHeaders(1).Tag <> "" Then
            It.Text = Format(Rs.Fields(0), lw1.ColumnHeaders(1).Tag)
        Else
            It.Text = Rs.Fields(0)
        End If
        'El resto de cmpos
        For NumRegElim = 2 To CInt(RecuperaValor(lw1.Tag, 2))
            If IsNull(Rs.Fields(NumRegElim - 1)) Then
                It.SubItems(NumRegElim - 1) = " "
            Else
                If lw1.ColumnHeaders(NumRegElim).Tag <> "" Then
                    It.SubItems(NumRegElim - 1) = Format(Rs.Fields(NumRegElim - 1), lw1.ColumnHeaders(NumRegElim).Tag)
                Else
                    It.SubItems(NumRegElim - 1) = Rs.Fields(NumRegElim - 1)
                End If
            End If
        Next
        It.SmallIcon = ElIcono
        Rs.MoveNext
    Wend
    Rs.Close
    Set Rs = Nothing
    
    
    Exit Sub
ECargaDatosLW:
    MuestraError Err.Number
    Set Rs = Nothing
    
End Sub


Private Sub lw1_DblClick()
Dim Seleccionado As Long
    
    If Modo <> 2 Then Exit Sub
    If lw1.ListItems.Count = 0 Then Exit Sub
    If lw1.SelectedItem Is Nothing Then Exit Sub

    If Me.DatosADevolverBusqueda <> "" Then
        'De momento NO dejo continuar
        MsgBox "Esta buscando un proveedor. No puede ver los documentos.", vbExclamation
        Exit Sub
    End If
    
    Screen.MousePointer = vbHourglass
    
    'Llegados aqui
    Select Case CByte(RecuperaValor(lw1.Tag, 1))
    Case 0
        'Pedidos de compra
'        Set frmPed = New frmComEntPedidos
'        frmPed.Pedido = lw1.SelectedItem.Text ' .SelectedItem.SubItems(1)
'        frmPed.Show vbModal
'        Set frmPed = Nothing
        
    Case 1
        'Albaranes de compra
        Set frmAlb = New frmComEntAlbaranes
        frmAlb.hcoCodProve = Data1.Recordset!CodProve ' .SelectedItem.SubItems(1)
        frmAlb.hcoCodMovim = lw1.SelectedItem.Text
        frmAlb.hcoFechaMovim = lw1.SelectedItem.SubItems(1)
        frmAlb.Show vbModal
        Set frmAlb = Nothing
    
    Case 2
        'Facturas
        Set frmFac = New frmComHcoFacturas
        frmFac.hcoCodProve = Data1.Recordset!CodProve ' .SelectedItem.SubItems(1)
        frmFac.hcoCodMovim = lw1.SelectedItem.Text
        frmFac.Factura = lw1.SelectedItem.Text
        frmFac.hcoFechaMovim = lw1.SelectedItem.SubItems(2)
        frmFac.Show vbModal
        Set frmFac = Nothing
        
    End Select
        
    'Pase lo que pase, por si acaso, cargamos el lw
    lw1.SetFocus
    Seleccionado = lw1.SelectedItem.Index
    CargaDatosLW
    lw1.SelectedItem.Selected = False
    Set lw1.SelectedItem = Nothing
    If lw1.ListItems.Count >= Seleccionado Then
            lw1.ListItems(Seleccionado).Selected = True
            lw1.ListItems(Seleccionado).EnsureVisible
    End If
End Sub

Private Sub Modificar()
'Modifica registre en les taules de Llínies
Dim nomFrame As String
Dim V As Integer
Dim Cad As String
Dim SQL As String
Dim b  As Boolean

    On Error GoTo eModificar

    Conn.BeginTrans
    
    b = True
    If Trim(Text1(1).Text) <> Trim(NombreAnt) Or Trim(Text1(3).Text) <> Trim(DomAnt) Or _
        Trim(Text1(4).Text) <> Trim(cPostalAnt) Or Trim(Text1(5).Text) <> Trim(PobAnt) Or _
        Trim(Text1(6).Text) <> Trim(ProAnt) Or Trim(Text1(7).Text) <> Trim(NifAnt) Or _
        Trim(Text1(21).Text) <> Trim(TelAnt) Then
        b = ModificarAlbaranes()
    End If
        
    ' modificamos los datos del campo
    If b Then
        If ModificaDesdeFormulario2(Me, 1) Then
            
            ModificarDatosCuentaContable

'            Or Trim(Text1(29).Text) <> Trim(IbanAnt) Or _
'        Trim(Text1(15).Text) <> Trim(BancoAnt) Or Trim(Text1(16).Text) <> Trim(SucurAnt) Or _
'        Trim(Text1(17).Text) <> Trim(DigitoAnt) Or Trim(Text1(18).Text) <> Trim(CuentaAnt) Or _
'        Trim(Text1(13).Text) <> Trim(ForpaAnt)
            TerminaBloquear
            PosicionarData
        End If
    
        Conn.CommitTrans
        Exit Sub
    End If
    

eModificar:
    Conn.RollbackTrans
    MuestraError Err.Number, "Modificando lineas"

End Sub

Public Function ModificarAlbaranes() As Boolean
'Comprobar si hay registros a Mostrar antes de abrir el Informe
Dim SQL As String
Dim Rs As ADODB.Recordset

    On Error GoTo eModificarEntradas

    ModificarAlbaranes = False

    SQL = "update scaalp "
    SQL = SQL & " set nomprove = " & DBSet(Text1(1).Text, "T")
    SQL = SQL & ", domprove = " & DBSet(Text1(3).Text, "T")
    SQL = SQL & ", codpobla = " & DBSet(Text1(4).Text, "T")
    SQL = SQL & ", pobprove = " & DBSet(Text1(5).Text, "T")
    SQL = SQL & ", proprove = " & DBSet(Text1(6).Text, "T")
    SQL = SQL & ", nifprove = " & DBSet(Text1(7).Text, "T")
    SQL = SQL & ", telprove = " & DBSet(Text1(21).Text, "T")
    SQL = SQL & "  where codprove = " & DBSet(Text1(0).Text, "N")
    
    Conn.Execute SQL
    
    ModificarAlbaranes = True
    Exit Function
    
eModificarEntradas:
    MuestraError Err.Number, "Modificar Datos Proveedor Albaranes Compra", Err.Description
End Function


Private Sub ModificarDatosCuentaContable()
Dim SQL As String
Dim Cad As String

    On Error GoTo eModificarDatosCuentaContable


    If Text1(1).Text <> NombreAnt Or Text1(15).Text <> BancoAnt Or Text1(16).Text <> SucurAnt Or Text1(17).Text <> DigitoAnt Or Text1(18).Text <> CuentaAnt Or _
       DomAnt <> Text1(3).Text Or cPostalAnt <> Text1(4).Text Or PobAnt <> Text1(5).Text Or ProAnt <> Text1(6).Text Or NifAnt <> Text1(7).Text Or _
       EMaiAnt <> Text1(14).Text Or WebAnt <> Text1(10).Text Or _
       IbanAnt <> Text1(29).Text Or ForpaAnt <> Text1(13).Text Then
        
        Cad = "Se han producido cambios en datos del Proveedor. " '& vbCrLf
        
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
        
        Cad = Cad & vbCrLf & vbCrLf & "¿ Desea actualizarlos en la Contabilidad ?" & vbCrLf & vbCrLf
        
        If MsgBox(Cad, vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
                        
            SQL = "update cuentas set nommacta = " & DBSet(Trim(Text1(1).Text), "T")
            SQL = SQL & ", razosoci = " & DBSet(Trim(Text1(2).Text), "T")
            SQL = SQL & ", dirdatos = " & DBSet(Trim(Text1(3).Text), "T")
            SQL = SQL & ", codposta = " & DBSet(Trim(Text1(4).Text), "T")
            SQL = SQL & ", despobla = " & DBSet(Trim(Text1(5).Text), "T")
            SQL = SQL & ", desprovi = " & DBSet(Trim(Text1(6).Text), "T")
            SQL = SQL & ", nifdatos = " & DBSet(Trim(Text1(7).Text), "T")
            SQL = SQL & ", maidatos = " & DBSet(Trim(Text1(14).Text), "T")
            SQL = SQL & ", webdatos = " & DBSet(Trim(Text1(10).Text), "T")
            SQL = SQL & ", forpa = " & DBSet(Trim(Text1(17).Text), "N")
            
            If vParamAplic.ContabilidadNueva Then
                Dim vIban As String
                
                vIban = MiFormat(Text1(29).Text, "") & MiFormat(Text1(15).Text, "0000") & MiFormat(Text1(16).Text, "0000") & MiFormat(Text1(17).Text, "00") & MiFormat(Text1(18).Text, "0000000000")
            
                SQL = SQL & ", iban = " & DBSet(vIban, "T")
                SQL = SQL & ", codpais = 'ES' "
                
            Else
                SQL = SQL & ", entidad = " & DBSet(Trim(Text1(15).Text), "T", "S")
                SQL = SQL & ", oficina = " & DBSet(Trim(Text1(16).Text), "T", "S")
                SQL = SQL & ", cc = " & DBSet(Trim(Text1(17).Text), "T", "S")
                SQL = SQL & ", cuentaba = " & DBSet(Trim(Text1(18).Text), "T", "S")
                
                '[Monica]22/11/2013: tema iban
                If vEmpresa.HayNorma19_34Nueva = 1 Then
                    SQL = SQL & ", iban = " & DBSet(Trim(Text1(29).Text), "T", "S")
                End If
            End If
            SQL = SQL & " where codmacta = " & DBSet(Trim(Text1(9).Text), "T")
                        
            ConnConta.Execute SQL
                        
'            MsgBox "Datos de Cuenta modificados correctamente.", vbExclamation
                        
        End If
    End If
    
    Exit Sub
    
eModificarDatosCuentaContable:
    MuestraError Err.Number, "Modificar Datos Cuenta Contable", Err.Description
End Sub


