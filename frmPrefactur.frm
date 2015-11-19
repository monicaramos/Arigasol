VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmPrefactur 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Informe Prefacturación"
   ClientHeight    =   6975
   ClientLeft      =   45
   ClientTop       =   2430
   ClientWidth     =   6825
   Icon            =   "frmPrefactur.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6975
   ScaleWidth      =   6825
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog cd1 
      Left            =   7680
      Top             =   5160
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame FrameCobros 
      Height          =   6975
      Left            =   30
      TabIndex        =   13
      Top             =   -30
      Width           =   6795
      Begin VB.Frame Frame4 
         Caption         =   "Tipo Facturación"
         ForeColor       =   &H00972E0B&
         Height          =   675
         Left            =   570
         TabIndex        =   31
         Top             =   4650
         Width           =   5265
         Begin VB.OptionButton Option1 
            Caption         =   "Tarjetas"
            Height          =   255
            Index           =   3
            Left            =   1410
            TabIndex        =   35
            Top             =   240
            Width           =   885
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Clientes"
            Height          =   255
            Index           =   2
            Left            =   270
            TabIndex        =   34
            Top             =   240
            Value           =   -1  'True
            Width           =   915
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Clientes/Tarjetas"
            Height          =   255
            Index           =   4
            Left            =   2490
            TabIndex        =   33
            Top             =   240
            Width           =   1575
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Interna"
            Height          =   255
            Index           =   5
            Left            =   4200
            TabIndex        =   32
            Top             =   240
            Width           =   915
         End
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   3180
         TabIndex        =   9
         Text            =   "Combo2"
         Top             =   4290
         Width           =   1275
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   4530
         TabIndex        =   10
         Text            =   "Combo1"
         Top             =   4290
         Width           =   1365
      End
      Begin VB.Frame Frame1 
         Caption         =   "Clasificado por"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00972E0B&
         Height          =   675
         Left            =   570
         TabIndex        =   40
         Top             =   4650
         Width           =   3045
         Begin VB.OptionButton Option1 
            Caption         =   "Cliente"
            Height          =   225
            Index           =   1
            Left            =   1920
            TabIndex        =   42
            Top             =   270
            Width           =   1035
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Forma de Pago"
            Height          =   225
            Index           =   0
            Left            =   330
            TabIndex        =   41
            Top             =   270
            Width           =   1455
         End
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Salta página por Cliente"
         Height          =   195
         Left            =   3810
         TabIndex        =   39
         Top             =   4920
         Width           =   1995
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   7
         Left            =   1860
         MaxLength       =   3
         TabIndex        =   5
         Top             =   2880
         Width           =   830
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   6
         Left            =   1860
         MaxLength       =   3
         TabIndex        =   4
         Top             =   2520
         Width           =   830
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   6
         Left            =   2745
         Locked          =   -1  'True
         TabIndex        =   30
         Text            =   "Text5"
         Top             =   2520
         Width           =   3165
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   7
         Left            =   2745
         Locked          =   -1  'True
         TabIndex        =   29
         Text            =   "Text5"
         Top             =   2895
         Width           =   3135
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Simulación de Facturación"
         Height          =   195
         Left            =   600
         TabIndex        =   8
         Top             =   4260
         Width           =   2175
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   5
         Left            =   2745
         Locked          =   -1  'True
         TabIndex        =   23
         Text            =   "Text5"
         Top             =   3795
         Width           =   3135
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   4
         Left            =   2745
         Locked          =   -1  'True
         TabIndex        =   22
         Text            =   "Text5"
         Top             =   3420
         Width           =   3135
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   5
         Left            =   1845
         MaxLength       =   2
         TabIndex        =   7
         Top             =   3795
         Width           =   830
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   4
         Left            =   1845
         MaxLength       =   2
         TabIndex        =   6
         Top             =   3420
         Width           =   830
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   3
         Left            =   1845
         MaxLength       =   10
         TabIndex        =   3
         Tag             =   "Código Postal|T|S|||clientes|codposta|||"
         Top             =   1980
         Width           =   1050
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   2
         Left            =   1845
         MaxLength       =   10
         TabIndex        =   2
         Tag             =   "Código Postal|T|S|||clientes|codposta|||"
         Top             =   1620
         Width           =   1050
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   5250
         TabIndex        =   12
         Top             =   6270
         Width           =   975
      End
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   4050
         TabIndex        =   11
         Top             =   6270
         Width           =   975
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   0
         Left            =   1845
         MaxLength       =   6
         TabIndex        =   0
         Top             =   660
         Width           =   830
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   1
         Left            =   1845
         MaxLength       =   6
         TabIndex        =   1
         Top             =   1035
         Width           =   830
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   2745
         Locked          =   -1  'True
         TabIndex        =   15
         Text            =   "Text5"
         Top             =   660
         Width           =   3135
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   2745
         Locked          =   -1  'True
         TabIndex        =   14
         Text            =   "Text5"
         Top             =   1035
         Width           =   3135
      End
      Begin MSComctlLib.ProgressBar Pb1 
         Height          =   255
         Left            =   510
         TabIndex        =   27
         Top             =   5940
         Visible         =   0   'False
         Width           =   5670
         _ExtentX        =   10001
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Factura"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   8
         Left            =   3180
         TabIndex        =   44
         Top             =   4080
         Width           =   960
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Cliente"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   7
         Left            =   4530
         TabIndex        =   43
         Top             =   4080
         Width           =   840
      End
      Begin VB.Label Label4 
         Caption         =   "Desde"
         Height          =   195
         Index           =   6
         Left            =   960
         TabIndex        =   38
         Top             =   2520
         Width           =   465
      End
      Begin VB.Label Label4 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   5
         Left            =   960
         TabIndex        =   37
         Top             =   2895
         Width           =   420
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Colectivo"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   4
         Left            =   570
         TabIndex        =   36
         Top             =   2280
         Width           =   660
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   3
         Left            =   1545
         MouseIcon       =   "frmPrefactur.frx":000C
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar colectivo"
         Top             =   2880
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   2
         Left            =   1545
         MouseIcon       =   "frmPrefactur.frx":015E
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar colectivo"
         Top             =   2535
         Width           =   240
      End
      Begin VB.Label Label4 
         Caption         =   "Recalculando Albaranes en temporal:"
         Height          =   195
         Index           =   3
         Left            =   540
         TabIndex        =   28
         Top             =   6270
         Visible         =   0   'False
         Width           =   3435
      End
      Begin VB.Image imgAyuda 
         Height          =   240
         Index           =   0
         Left            =   2820
         MousePointer    =   4  'Icon
         Tag             =   "-1"
         ToolTipText     =   "Ayuda"
         Top             =   4290
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   5
         Left            =   1545
         MouseIcon       =   "frmPrefactur.frx":02B0
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar F.Pago"
         Top             =   3795
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   4
         Left            =   1545
         MouseIcon       =   "frmPrefactur.frx":0402
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar F.Pago"
         Top             =   3420
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Forma de Pago"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   2
         Left            =   600
         TabIndex        =   26
         Top             =   3180
         Width           =   1080
      End
      Begin VB.Label Label4 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   1
         Left            =   960
         TabIndex        =   25
         Top             =   3795
         Width           =   420
      End
      Begin VB.Label Label4 
         Caption         =   "Desde"
         Height          =   195
         Index           =   0
         Left            =   960
         TabIndex        =   24
         Top             =   3420
         Width           =   465
      End
      Begin VB.Label Label4 
         Caption         =   "Fecha albarán"
         ForeColor       =   &H00972E0B&
         Height          =   255
         Index           =   16
         Left            =   600
         TabIndex        =   21
         Top             =   1380
         Width           =   1815
      End
      Begin VB.Label Label4 
         Caption         =   "Desde"
         Height          =   195
         Index           =   15
         Left            =   960
         TabIndex        =   20
         Top             =   1620
         Width           =   465
      End
      Begin VB.Label Label4 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   14
         Left            =   960
         TabIndex        =   19
         Top             =   1980
         Width           =   420
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   2
         Left            =   1545
         Picture         =   "frmPrefactur.frx":0554
         ToolTipText     =   "Buscar fecha"
         Top             =   1620
         Width           =   240
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   3
         Left            =   1545
         Picture         =   "frmPrefactur.frx":05DF
         ToolTipText     =   "Buscar fecha"
         Top             =   1980
         Width           =   240
      End
      Begin VB.Label Label4 
         Caption         =   "Desde"
         Height          =   195
         Index           =   13
         Left            =   960
         TabIndex        =   18
         Top             =   660
         Width           =   465
      End
      Begin VB.Label Label4 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   12
         Left            =   960
         TabIndex        =   17
         Top             =   1035
         Width           =   420
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Cliente"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   11
         Left            =   600
         TabIndex        =   16
         Top             =   420
         Width           =   495
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   0
         Left            =   1545
         MouseIcon       =   "frmPrefactur.frx":066A
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar Cliente"
         Top             =   660
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   1
         Left            =   1545
         MouseIcon       =   "frmPrefactur.frx":07BC
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar Cliente"
         Top             =   1035
         Width           =   240
      End
   End
End
Attribute VB_Name = "frmPrefactur"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' +-+-+-+-+-+-+-+-+-+-+-
' +-+- Autor MANOLO +-+-
' +-+-+-+-+-+-+-+-+-+-+-

Option Explicit

Public NumCod As String 'Para indicar cod. Traspaso,Movimiento, etc. que llama
                        'Para indicar nº oferta a imprimir

Public CadTag As String 'Cadena con el Tag del campo que se va a poner en D/H en los listados
                        'Se necesita si el tipo de codigo es texto

    
Private Conexion As Byte
'1.- Conexión a BD Ariges  2.- Conexión a BD Conta

Private HaDevueltoDatos As Boolean

Private WithEvents frmFpa As frmManFpago 'F.Pago
Attribute frmFpa.VB_VarHelpID = -1
Private WithEvents frmcli As frmManClien 'Clientes
Attribute frmcli.VB_VarHelpID = -1
Private WithEvents frmC As frmCal 'calendario fechas
Attribute frmC.VB_VarHelpID = -1
Private WithEvents frmCol As frmManCoope 'Colectivo
Attribute frmCol.VB_VarHelpID = -1

'GENERALES PARA PASARLE A CRYSTAL REPORT
Private cadFormula As String 'Cadena con la FormulaSelection para Crystal Report
Private cadParam As String 'Cadena con los parametros para Crystal Report
Private numParam As Byte 'Numero de parametros que se pasan a Crystal Report
Private cadSelect As String 'Cadena para comprobar si hay datos antes de abrir Informe
Private cadTitulo As String 'Titulo para la ventana frmImprimir
Private cadNombreRPT As String 'Nombre del informe

Dim indCodigo As Integer 'indice para txtCodigo
Dim indFrame As Single 'nº de frame en el que estamos
 
'Se inicializan para cada Informe (tabla de BD a la que hace referencia
Dim tabla As String
Dim Codigo As String 'Código para FormulaSelection de Crystal Report
Dim TipCod As String
Dim Orden1 As String 'Campo de Ordenacion (por codigo) para Cristal Report
Dim Orden2 As String 'Campo de Ordenacion (por nombre) para Cristal Report

Dim PrimeraVez As Boolean

Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, 0, cerrar
    If cerrar Then Unload Me
End Sub

Private Sub Check2_Click()
    Frame1.Enabled = (Check2.Value = 0)
    Frame1.visible = (Check2.Value = 0)
    Check1.Enabled = (Check2.Value = 0)
    Check1.visible = (Check2.Value = 0)
    Frame4.Enabled = (Check2.Value = 1)
    Frame4.visible = (Check2.Value = 1)
End Sub

Private Sub Check2_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub cmdAceptar_Click()
    Screen.MousePointer = vbHourglass
    cmdAceptar1_Click
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmdAceptar1_Click()
Dim cDesde As String, cHasta As String 'cadena codigo Desde/Hasta
Dim nDesde As String, nHasta As String 'cadena Descripcion Desde/Hasta
Dim cadTABLA As String, cOrden As String
Dim I As Byte
Dim CadArtic As String

Dim Sql5 As String
Dim Sql3 As String

InicializarVbles
    
    
    '========= PARAMETROS  =============================
    'Añadir el parametro de Empresa
    cadParam = cadParam & "|pEmpresa=""" & vEmpresa.nomEmpre & """|"
    numParam = numParam + 1
    
    
    'D/H Cliente
    cDesde = Trim(txtCodigo(0).Text)
    cHasta = Trim(txtCodigo(1).Text)
    nDesde = txtNombre(0).Text
    nHasta = txtNombre(1).Text
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{" & tabla & ".codsocio}"
        TipCod = "N"
        If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHcliente= """) Then Exit Sub
    End If
    
    'D/H Fecha factura
    cDesde = Trim(txtCodigo(2).Text)
    cHasta = Trim(txtCodigo(3).Text)
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{" & tabla & ".fecalbar}"
        TipCod = "F"
        If Not PonerDesdeHasta(cDesde, cHasta, "", "", "pDHfechaFac= """) Then Exit Sub
    End If
    
    'D/H F.Pago
    cDesde = Trim(txtCodigo(4).Text)
    cHasta = Trim(txtCodigo(5).Text)
    nDesde = txtNombre(4).Text
    nHasta = txtNombre(5).Text
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{scaalb.codforpa}"
        TipCod = "N"
        If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDFpago= """) Then Exit Sub
    End If
    
    'D/H Colectivo
    cDesde = Trim(txtCodigo(6).Text)
    cHasta = Trim(txtCodigo(7).Text)
    nDesde = txtNombre(6).Text
    nHasta = txtNombre(7).Text
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{ssocio.codcoope}"
        TipCod = "N"
        If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDCoope= """) Then Exit Sub
    End If
    
    
    'Comprobar si hay registros a Mostrar antes de abrir el Informe
    cadTABLA = tabla & " INNER JOIN ssocio ON " & tabla & ".codsocio=ssocio.codsocio "
    
    '[Monica]19/06/2013: gasoleo b
    Sql3 = ""
    
    If vParamAplic.Cooperativa = 1 Or vParamAplic.Cooperativa = 2 Then
        cadTABLA = "(" & cadTABLA & ") INNER JOIN sartic ON scaalb.codartic = sartic.codartic "
    
        Select Case Combo2.ListIndex
            Case 0
                Sql3 = "not scaalb.codartic in (select codartic from sartic where tipogaso = 3 union " & _
                                                         "select if(artdto is null, -1, artdto) from sartic where tipogaso = 3)"
            Case 1
                Sql3 = "scaalb.codartic in (select codartic from sartic where tipogaso = 3 And esdomiciliado = 0 union " & _
                                                     "select if(artdto is null, -1, artdto) from sartic where tipogaso = 3 And esdomiciliado = 0)"
            Case 2
                Sql3 = "scaalb.codartic in (select codartic from sartic where tipogaso = 3 And esdomiciliado = 1 union " & _
                                                     "select if(artdto is null, -1, artdto) from sartic where tipogaso = 3 And esdomiciliado = 1)"
        End Select
        
        If Not AnyadirAFormula(cadSelect, Sql3) Then Exit Sub
    End If
    
    
    '[Monica]24/06/2013: cargo una temporal con los articulos
    CadArtic = CadenaArticulos(Sql3)
    
    
    
    
    '[Monica]18/01/2013: condicion de tipo de socios
    Select Case Combo1.ListIndex
        Case 0
            
        Case 1 ' clientes con bonificacion especial
            If Not AnyadirAFormula(cadSelect, "ssocio.bonifesp = 1") Then Exit Sub
            If Not AnyadirAFormula(cadFormula, "{ssocio.bonifesp} = 1") Then Exit Sub
        
        Case 2 ' clientes sin bonificacion especial
            If Not AnyadirAFormula(cadSelect, "ssocio.bonifesp = 0") Then Exit Sub
            If Not AnyadirAFormula(cadFormula, "{ssocio.bonifesp} = 0") Then Exit Sub
    End Select
    
    
    If Check2.Value Then
        If Not AnyadirAFormula(cadSelect, "scaalb.numfactu = 0 and scaalb.codforpa <> 98") Then Exit Sub
        
        '[Monica]25/01/2013: Comprobamos que no hay ningun albaran de importe superior a 2500
        If vParamAplic.Cooperativa = 1 And vParamAplic.LimiteFra <> 0 Then
            Sql5 = "select count(*) from " & cadTABLA & " where " & cadSelect & " and scaalb.importel > " & DBSet(vParamAplic.LimiteFra, "N")
            If TotalRegistros(Sql5) > 0 Then
                MsgBox "Hay albaranes con importe superior a " & vParamAplic.LimiteFra & " euros. Revise.", vbExclamation
                Exit Sub
            End If
        End If
    End If
    
    
    If HayRegParaInforme(cadTABLA, cadSelect) Then
        If Check2.Value Then
            If ProcesoCargaTemporal(cadTABLA, cadSelect) Then
                cadTitulo = "Informe Prefacturación por Cliente"
                cadNombreRPT = "rPrefacturClienteBon.rpt"
            
                cadParam = cadParam & "pSalto=" & Me.Check1.Value & "|"
                numParam = numParam + 1
                  
                cadFormula = "{tmpscaalb.codusu} = " & vSesion.Codigo
            End If
        Else
        
            If Not AnyadirAFormula(cadFormula, "{tmpinformes.codusu} = " & vSesion.Codigo) Then Exit Sub
        
            If Option1(0).Value Then
              cadTitulo = "Informe Prefacturación por Forma de Pago"
              cadNombreRPT = "rPrefactur.rpt"
            Else
              cadTitulo = "Informe Prefacturación por Cliente"
              cadNombreRPT = "rPrefacturCliente.rpt"
          
              cadParam = cadParam & "pSalto=" & Me.Check1.Value & "|"
              numParam = numParam + 1
            End If
        End If
        LlamarImprimir
          'AbrirVisReport
    End If

    Screen.MousePointer = vbDefault


End Sub

Private Function ProcesoCargaTemporal(cadTABLA As String, cadSelect As String) As Boolean
' El proceso tiene varios pasos:
' 1 - Cargamos todos los albaranes de la prefacturacion en la tabla temporal tmpscaalb
' 2 - Procesamos los registros de tmpscaalb recalculando importes con los nuevos precios de bonificacion
' 3 - Sobre estos albaranes ejecutamos la simulacion de una facturacion, guardando nrofactura y lineas de bonificacion
Dim sql As String
Dim Sql2 As String
Dim Sql3 As String
Dim Tipo As Byte
Dim b As Boolean

    On Error GoTo eProcesoCargaTemporal
    
    ProcesoCargaTemporal = False
    
    ' 1 - Cargamos la tabla sobre la que está el formulario para que vean y puedan cambiar precios de ultima compra
    sql = "delete from tmpscaalb where codusu = " & vSesion.Codigo
    Conn.Execute sql
    
    sql = "insert into tmpscaalb (codusu,codclave,codsocio,numtarje,numalbar,fecalbar,horalbar,codturno,"
    sql = sql & "codartic,cantidad,preciove,importel,codforpa,matricul,codtraba,numfactu,numlinea,declaradogp,"
    sql = sql & "precioinicial) select " & vSesion.Codigo & ",scaalb.codclave,scaalb.codsocio,scaalb.numtarje,scaalb.numalbar,scaalb.fecalbar,"
    sql = sql & "scaalb.horalbar,scaalb.codturno,scaalb.codartic,scaalb.cantidad,scaalb.preciove,scaalb.importel,scaalb.codforpa,scaalb.matricul,scaalb.codtraba,scaalb.numfactu,"
    sql = sql & "scaalb.numlinea,scaalb.declaradogp,scaalb.precioinicial"
    sql = sql & " from " & cadTABLA
    
    If cadSelect <> "" Then
        sql = sql & " where " & Replace(Replace(cadSelect, "{", ""), "}", "")
    End If
    
    Conn.Execute sql
    
    ' 2 - Proceso de recalculo de importes en al tabla temporal de albaranes
    sql = "select distinct " & vSesion.Codigo & ", tmpscaalb.codartic, sartic.nomartic, sartic.ultpreci "
    Sql2 = "select tmpscaalb.codclave "
    sql = sql & " from ((((tmpscaalb INNER JOIN sartic ON tmpscaalb.codartic = sartic.codartic) "
    Sql2 = Sql2 & " from ((((tmpscaalb INNER JOIN sartic ON tmpscaalb.codartic = sartic.codartic) "
    sql = sql & " INNER JOIN sfamia ON sartic.codfamia = sfamia.codfamia and sfamia.tipfamia = 1) " ' solo articulos de la familia de carburantes
    Sql2 = Sql2 & " INNER JOIN sfamia ON sartic.codfamia = sfamia.codfamia and sfamia.tipfamia = 1) "
    sql = sql & " INNER JOIN ssocio ON tmpscaalb.codsocio = ssocio.codsocio and ssocio.bonifesp = 1) " ' solo clientes que sean de bonificacion especial
    Sql2 = Sql2 & "  INNER JOIN ssocio ON tmpscaalb.codsocio = ssocio.codsocio and ssocio.bonifesp = 1) "
    '[Monica]04/01/2013: Efectivos
    sql = sql & " INNER JOIN sforpa ON tmpscaalb.codforpa = sforpa.codforpa and sforpa.tipforpa <> 0  and sforpa.tipforpa <> 6) " ' solo formas de pago que no sean de efectivo
    Sql2 = Sql2 & " INNER JOIN sforpa ON tmpscaalb.codforpa = sforpa.codforpa and sforpa.tipforpa <> 0  and sforpa.tipforpa <> 6) "
    '[Monica]28/12/2011:tenemos que saber que articulos tienen bonificacion
    sql = sql & " INNER JOIN smargen ON tmpscaalb.codsocio = smargen.codsocio and tmpscaalb.codartic = smargen.codartic " ' solo los articulos que tengan bonificacion
    Sql2 = Sql2 & " INNER JOIN smargen ON tmpscaalb.codsocio = smargen.codsocio and tmpscaalb.codartic = smargen.codartic "
    
    If cadSelect <> "" Then
        sql = sql & " where codusu = " & vSesion.Codigo
        Sql2 = Sql2 & " where codusu = " & vSesion.Codigo
    End If
    If TotalRegistrosConsulta(Sql2) <> 0 Then
        ' cargamos la tabla sobre la que está el formulario para que vean y puedan cambiar precios de ultima compra
        Sql3 = "delete from tmpinformes where codusu = " & vSesion.Codigo
        Conn.Execute Sql3
        
        Sql3 = "insert into tmpinformes(codusu, codigo1, nombre1, precio2) " & sql
        Conn.Execute Sql3
        
        frmPreciosArt.Show vbModal
        Label4(3).Caption = "Recalculando Albaranes en temporal:"
        b = ModificacionAlbaranes(Sql2, "tmpscaalb", Pb1, Label4(3))
    Else
        b = True
    End If
    
    If b Then
        ' 3 - Proceso de facturacion en donde cargamos en numfactu el nro de factura
        '     y calculamos la linea de bonificacion
        Tipo = 1
        If Option1(2).Value Then Tipo = 1
        If Option1(3).Value Then Tipo = 0
        If Option1(4).Value Then Tipo = 2
        
        If Option1(5).Value Then Tipo = 3

        If (Tipo = 2 Or Tipo = 1) And Combo2.ListIndex = 0 Then
            b = SimulacionFacturacion(1, Pb1, Label4(3), 0)
        End If
        If b Then
            If Tipo = 2 Or Tipo = 0 Or Tipo = 3 Then
                If Tipo = 3 Then
                    b = SimulacionFacturacion(Tipo, Pb1, Label4(3), Combo2.ListIndex) '[Monica]combo2.listindex era antes 0
                Else
                    b = SimulacionFacturacion(0, Pb1, Label4(3), Combo2.ListIndex)
                End If
            End If
        End If
                
        If b Then
            b = Borrado
        End If
                
        ProcesoCargaTemporal = b
        Exit Function
    End If

eProcesoCargaTemporal:
    MuestraError Err.Number, "Cargando Tabla Temporal", Err.Description
End Function


Private Function Borrado() As Boolean
Dim sql As String

    On Error GoTo eBorrado

    Borrado = False

    '[Monica]23/07/2013: quito las que no tienen nro de factura
    sql = "delete from tmpscaalb where numfactu = 0 and codusu = " & vSesion.Codigo
    Conn.Execute sql

    Borrado = True
    Exit Function
    
eBorrado:

End Function


Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub



Private Sub Combo2_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub Combo2_LostFocus()
    Select Case Combo2.ListIndex
        Case 0
            Frame4.visible = True
            Frame4.Enabled = True
'[Monica]15/07/2013: añadido esto
            Option1(2).Enabled = True
            Option1(4).Enabled = True
            Option1(2).Value = True
            
        Case 1
'            Frame4.Enabled = False
'            Option1(3).Value = True
'[Monica]15/07/2013: cambiado por esto
            Frame4.visible = True
            Frame4.Enabled = True
            Option1(2).Enabled = False
            Option1(4).Enabled = False
            Option1(3).Value = True
        Case 2
'            Frame4.Enabled = False
'            Option1(3).Value = True
'[Monica]15/07/2013: cambiado por esto
            Frame4.visible = True
            Frame4.Enabled = True
            Option1(2).Enabled = False
            Option1(4).Enabled = False
            Option1(3).Value = True
    End Select

End Sub




Private Sub Form_Activate()
    If PrimeraVez Then
        PrimeraVez = False
        PonerFoco txtCodigo(0)
    End If
    Me.Combo2.ListIndex = 0
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
Dim h As Integer, w As Integer
Dim List As Collection
Dim I As Integer

    PrimeraVez = True
    limpiar Me

    'IMAGES para busqueda
    For I = 0 To 5
        Me.imgBuscar(I).Picture = frmPpal.imgListImages16.ListImages(1).Picture
    Next I
    
     imgAyuda(0).Picture = frmPpal.ImageListB.ListImages(10).Picture
    '###Descomentar
'    CommitConexion
         
    FrameCobrosVisible True, h, w
    indFrame = 5
    tabla = "scaalb"
            
    'Esto se consigue poneinedo el cancel en el opcion k corresponda
    Me.cmdCancel.Cancel = True
    
    Combo2.Enabled = (vParamAplic.Cooperativa = 1 Or vParamAplic.Cooperativa = 2)
    Combo2.visible = (vParamAplic.Cooperativa = 1 Or vParamAplic.Cooperativa = 2)
    Label4(8).visible = (vParamAplic.Cooperativa = 1 Or vParamAplic.Cooperativa = 2)
    
    CargaCombo
    
    Me.Combo1.ListIndex = 0
    
    Check2_Click
    
    Option1(0).Value = True
'   Me.Width = w + 70
'   Me.Height = h + 350
End Sub

Private Sub frmC_Selec(vFecha As Date)
 'Fecha
    txtCodigo(CByte(imgFec(2).Tag)).Text = Format(vFecha, "dd/MM/yyyy")
End Sub

Private Sub frmCli_DatoSeleccionado(CadenaSeleccion As String)
'Form de Consulta de Clientes
    txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000000")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmCol_DatoSeleccionado(CadenaSeleccion As String)
'Form de Consulta de Colectivos
    txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmFpa_DatoSeleccionado(CadenaSeleccion As String)
'Form de Consulta de Colectivos
    txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "00")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub


Private Sub imgAyuda_Click(Index As Integer)
Dim vCadena As String
    Select Case Index
        Case 0
           ' "____________________________________________________________"
            vCadena = "Si está marcado simulamos el proceso de facturación de albaranes que" & vbCrLf & _
                      "no vengan del TPV, mediante los siguientes pasos : " & vbCrLf & vbCrLf & _
                      "     1- Prefacturación Bonificación: Se recalcula el importe del albaran" & vbCrLf & _
                      "con el precio introducido al que se le aplica el margen del cliente, si es" & vbCrLf & _
                      "de bonificación especial, y sólo si la forma de pago del albarán no es " & vbCrLf & _
                      "efectivo sobre artículos que sean de la familia combustible." & vbCrLf & _
                      "     2- Se calculan las líneas de bonificación si la tiene. " & vbCrLf & _
                      "     El resultado sale agrupado por cliente y factura. " & vbCrLf & vbCrLf & vbCrLf & _
                      "En caso contrario, saca un listado de los albaranes seleccionados" & vbCrLf & _
                      "ordenados de la manera indicada. " & vbCrLf & vbCrLf
                      
    End Select
    MsgBox vCadena, vbInformation, "Descripción de Ayuda"
    
End Sub

Private Sub imgFec_Click(Index As Integer)
'FEchas
    Dim esq, dalt As Long
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
       
    ' es desplega dalt i cap a la esquerra
    frmC.Left = esq + imgFec(Index).Parent.Left + 30
    frmC.Top = dalt + imgFec(Index).Parent.Top + imgFec(Index).Height + 420 + 30

    ' ***canviar l'index de imgFec pel 1r index de les imagens de buscar data***
    imgFec(2).Tag = Index 'independentment de les dates que tinga, sempre pose l'index en la 27
    If txtCodigo(Index).Text <> "" Then frmC.NovaData = txtCodigo(Index).Text

    frmC.Show vbModal
    Set frmC = Nothing
    PonerFoco txtCodigo(CByte(imgFec(2).Tag) + 2)
    ' ***************************
End Sub

Private Sub ListView1_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub imgBuscar_Click(Index As Integer)
   Select Case Index
        Case 0, 1 'CLIENTE
            AbrirFrmClientes (Index)
        
        Case 2, 3 'COLECTIVO
            AbrirFrmColectivo (Index + 4)
        
        Case 4, 5 'F.PAGO
            AbrirFrmFpagos (Index)
        
    End Select
    PonerFoco txtCodigo(indCodigo)
End Sub

Private Sub Optcodigo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then 'ENTER
        PonerFocoBtn Me.cmdAceptar
    ElseIf KeyAscii = 27 Then Unload Me  'ESC
    End If
End Sub

Private Sub OptNombre_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then 'ENTER
        PonerFocoBtn Me.cmdAceptar
    ElseIf KeyAscii = 27 Then Unload Me  'ESC
    End If
End Sub

Private Sub Option1_Click(Index As Integer)
    Check1.Enabled = (Option1(1).Value = True)
    If Option1(0).Value Then
        Check1.Value = 0
    End If
End Sub

Private Sub txtCodigo_GotFocus(Index As Integer)
    ConseguirFoco txtCodigo(Index), 3
End Sub

Private Sub txtCodigo_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    KEYdown KeyCode
End Sub

Private Sub txtCodigo_KeyPress(Index As Integer, KeyAscii As Integer)
'15/02/2007
'    KEYpress KeyAscii
    If KeyAscii = teclaBuscar Then
        Select Case Index
            Case 0: KEYBusqueda KeyAscii, 0 'cliente desde
            Case 1: KEYBusqueda KeyAscii, 1 'cliente hasta
            Case 4: KEYBusqueda KeyAscii, 4 'forma de pago desde
            Case 5: KEYBusqueda KeyAscii, 5 'forma de pago hasta
            Case 2: KEYFecha KeyAscii, 2 'fecha albaran desde
            Case 3: KEYFecha KeyAscii, 3 'fecha albaran hasta
        End Select
    Else
        KEYpress KeyAscii
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

Private Sub txtCodigo_LostFocus(Index As Integer)
Dim cad As String, cadTipo As String 'tipo cliente

    'Quitar espacios en blanco por los lados
    txtCodigo(Index).Text = Trim(txtCodigo(Index).Text)
    
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub

    
    Select Case Index
            
        Case 0, 1 'CLIENTE
            txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), "ssocio", "nomsocio", "codsocio", "N")
            If txtCodigo(Index).Text <> "" Then txtCodigo(Index).Text = Format(txtCodigo(Index).Text, "000000")
        
        Case 2, 3 'FECHAS
            If txtCodigo(Index).Text <> "" Then PonerFormatoFecha txtCodigo(Index)
            
        Case 6, 7 'COLECTIVO
            txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), "scoope", "nomcoope", "codcoope", "N")
            If txtCodigo(Index).Text <> "" Then txtCodigo(Index).Text = Format(txtCodigo(Index).Text, "000")
            
        Case 4, 5 'F.PAGO
            txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), "sforpa", "nomforpa", "codforpa", "N")
            If txtCodigo(Index).Text <> "" Then txtCodigo(Index).Text = Format(txtCodigo(Index).Text, "00")
    End Select
End Sub

Private Sub FrameCobrosVisible(visible As Boolean, ByRef h As Integer, ByRef w As Integer)
    Me.FrameCobros.visible = visible
'    If visible = True Then
'        Me.FrameCobros.Top = -90
'        Me.FrameCobros.Left = 0
'        Me.FrameCobros.Height = 6015
'        Me.FrameCobros.Width = 6555
'        w = Me.FrameCobros.Width
'        h = Me.FrameCobros.Height
'    End If
End Sub

Private Sub InicializarVbles()
    cadFormula = ""
    cadSelect = ""
    cadParam = ""
    numParam = 0
End Sub

Private Function PonerDesdeHasta(codD As String, codH As String, nomD As String, nomH As String, param As String) As Boolean
'IN: codD,codH --> codigo Desde/Hasta
'    nomD,nomH --> Descripcion Desde/Hasta
'Añade a cadFormula y cadSelect la cadena de seleccion:
'       "(codigo>=codD AND codigo<=codH)"
' y añade a cadParam la cadena para mostrar en la cabecera informe:
'       "codigo: Desde codD-nomd Hasta: codH-nomH"
Dim devuelve As String
Dim devuelve2 As String

    PonerDesdeHasta = False
    devuelve = CadenaDesdeHasta(codD, codH, Codigo, TipCod)
    If devuelve = "Error" Then Exit Function
    If Not AnyadirAFormula(cadFormula, devuelve) Then Exit Function
    If TipCod <> "F" Then 'Fecha
        If Not AnyadirAFormula(cadSelect, devuelve) Then Exit Function
    Else
        devuelve2 = CadenaDesdeHastaBD(codD, codH, Codigo, TipCod)
        If devuelve2 = "Error" Then Exit Function
        If Not AnyadirAFormula(cadSelect, devuelve2) Then Exit Function
    End If
    If devuelve <> "" Then
        If param <> "" Then
            'Parametro Desde/Hasta
            cadParam = cadParam & AnyadirParametroDH(param, codD, codH, nomD, nomH)
            numParam = numParam + 1
        End If
        PonerDesdeHasta = True
    End If
End Function

Private Sub LlamarImprimir()
    With frmImprimir
        .FormulaSeleccion = cadFormula
        .OtrosParametros = cadParam
        .NumeroParametros = numParam
        .SoloImprimir = False
        .Titulo = cadTitulo
        .NombreRPT = cadNombreRPT
        .ConSubInforme = True
        .Opcion = 0
        .Show vbModal
    End With
End Sub

Private Sub AbrirFrmClientes(indice As Integer)
    indCodigo = indice
    Set frmcli = New frmManClien
    frmcli.DatosADevolverBusqueda = "0|1|"
    frmcli.DeConsulta = True
    frmcli.CodigoActual = txtCodigo(indCodigo)
    frmcli.Show vbModal
    Set frmcli = Nothing
End Sub

Private Sub AbrirFrmFpagos(indice As Integer)
    indCodigo = indice
    Set frmFpa = New frmManFpago
    frmFpa.DatosADevolverBusqueda = "0|1|"
    frmFpa.DeConsulta = True
    frmFpa.CodigoActual = txtCodigo(indCodigo)
    frmFpa.Show vbModal
    Set frmFpa = Nothing
End Sub
 
Private Sub AbrirFrmColectivo(indice As Integer)
    indCodigo = indice
    Set frmCol = New frmManCoope
    frmCol.DatosADevolverBusqueda = "0|1|"
    frmCol.DeConsulta = True
    frmCol.CodigoActual = txtCodigo(indCodigo)
    frmCol.Show vbModal
    Set frmCol = Nothing
End Sub

 
 
Private Sub AbrirVisReport()
    Screen.MousePointer = vbHourglass
    CadenaDesdeOtroForm = ""
    With frmVisReport
        .FormulaSeleccion = cadFormula
'        .SoloImprimir = (Me.OptVisualizar(indFrame).Value = 1)
        .OtrosParametros = cadParam
        .NumeroParametros = numParam
        '##descomen
'        .MostrarTree = MostrarTree
'        .Informe = MIPATH & Nombre
'        .InfConta = InfConta
        '##
        
'        If NombreSubRptConta <> "" Then
'            .SubInformeConta = NombreSubRptConta
'        Else
'            .SubInformeConta = ""
'        End If
        '##descomen
'        .ConSubInforme = ConSubInforme
        '##
        .Opcion = ""
'        .ExportarPDF = (chkEMAIL.Value = 1)
        .Show vbModal
    End With
    
'    If Me.chkEMAIL.Value = 1 Then
'    '####Descomentar
'        If CadenaDesdeOtroForm <> "" Then frmEMail.Show vbModal
'    End If
    Unload Me
End Sub

Private Sub AbrirEMail()
    If CadenaDesdeOtroForm <> "" Then frmEMail.Show vbModal
End Sub

Private Sub CargaCombo()
    
    Combo1.Clear
    'Conceptos
    Combo1.AddItem "Todos"
    Combo1.ItemData(Combo1.NewIndex) = 0
    
    Combo1.AddItem "Con Bonif.Esp."
    Combo1.ItemData(Combo1.NewIndex) = 1
    
    Combo1.AddItem "Sin Bonif."
    Combo1.ItemData(Combo1.NewIndex) = 2


    Combo2.Clear
    'Tipos de facturacion
    Combo2.AddItem "Normal"
    Combo2.ItemData(Combo2.NewIndex) = 0
    
    Combo2.AddItem "Gas.B"
    Combo2.ItemData(Combo2.NewIndex) = 1
    
    ' 19/06/2013: solo en el caso de que sea Alzira hay gasoleo B domiciliado ( de momento )
    If vParamAplic.Cooperativa = 1 Then
        Combo2.AddItem "Gas.B Dom."
        Combo2.ItemData(Combo2.NewIndex) = 2
    End If
    
End Sub

Private Function CadenaArticulos(cadena As String) As Boolean
Dim sql As String
Dim Sql2 As String
Dim Rs As ADODB.Recordset

    On Error GoTo eCadenaArticulos

    CadenaArticulos = True
    
    sql = "delete from tmpinformes where codusu = " & vSesion.Codigo
    Conn.Execute sql

    If cadena = "" Then cadena = "(1=1)"

    sql = "select " & vSesion.Codigo & ", codartic from sartic where " & Replace(cadena, "scaalb", "sartic")

    sql = "insert into tmpinformes (codusu, importe1) " & sql
    
    Conn.Execute sql

    Exit Function
        
eCadenaArticulos:
    CadenaArticulos = False
    MuestraError Err.Number, "Cadena Articulos", Err.Description
End Function

        
        


