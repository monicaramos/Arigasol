VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmEstTarArt 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ventas Artículos por Tarjeta"
   ClientHeight    =   6495
   ClientLeft      =   45
   ClientTop       =   2430
   ClientWidth     =   6900
   Icon            =   "frmEstTarArt.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6495
   ScaleWidth      =   6900
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
      Height          =   6255
      Left            =   120
      TabIndex        =   13
      Top             =   120
      Width           =   6675
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Index           =   7
         Left            =   1860
         MaxLength       =   16
         TabIndex        =   6
         Tag             =   "Código de articulo|N|N|0|999999|sartic|codartic|000000|S|"
         Top             =   3480
         Width           =   1065
      End
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Index           =   8
         Left            =   1860
         MaxLength       =   16
         TabIndex        =   7
         Tag             =   "Código de articulo|N|N|0|999999|sartic|codartic|000000|S|"
         Top             =   3840
         Width           =   1065
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   7
         Left            =   2985
         Locked          =   -1  'True
         TabIndex        =   31
         Text            =   "Text5"
         Top             =   3480
         Width           =   3280
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   8
         Left            =   2985
         Locked          =   -1  'True
         TabIndex        =   30
         Text            =   "Text5"
         Top             =   3840
         Width           =   3280
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   6
         Left            =   1845
         MaxLength       =   2
         TabIndex        =   8
         Top             =   4320
         Width           =   830
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   6
         Left            =   2745
         Locked          =   -1  'True
         TabIndex        =   28
         Text            =   "Text5"
         Top             =   4320
         Width           =   3480
      End
      Begin VB.Frame Frame4 
         Caption         =   "Tipo Informe"
         ForeColor       =   &H00972E0B&
         Height          =   1000
         Left            =   960
         TabIndex        =   27
         Top             =   4920
         Width           =   2175
         Begin VB.OptionButton Option1 
            Caption         =   "Detalle"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   9
            Top             =   240
            Value           =   -1  'True
            Width           =   1815
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Resumen"
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   10
            Top             =   600
            Width           =   1815
         End
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   5
         Left            =   2745
         Locked          =   -1  'True
         TabIndex        =   26
         Text            =   "Text5"
         Top             =   2895
         Width           =   3495
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   4
         Left            =   2745
         Locked          =   -1  'True
         TabIndex        =   25
         Text            =   "Text5"
         Top             =   2520
         Width           =   3495
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   5
         Left            =   1845
         MaxLength       =   2
         TabIndex        =   5
         Top             =   2895
         Width           =   830
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   4
         Left            =   1845
         MaxLength       =   2
         TabIndex        =   4
         Top             =   2520
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
         Top             =   1920
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
         Top             =   1560
         Width           =   1050
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   4905
         TabIndex        =   12
         Top             =   5640
         Width           =   975
      End
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   3720
         TabIndex        =   11
         Top             =   5640
         Width           =   975
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   0
         Left            =   1860
         MaxLength       =   8
         TabIndex        =   0
         Top             =   600
         Width           =   830
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   1
         Left            =   1845
         MaxLength       =   8
         TabIndex        =   1
         Top             =   975
         Width           =   830
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   2745
         Locked          =   -1  'True
         TabIndex        =   23
         Text            =   "Text5"
         Top             =   600
         Width           =   3615
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   2745
         Locked          =   -1  'True
         TabIndex        =   24
         Text            =   "Text5"
         Top             =   975
         Width           =   3615
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   23
         Left            =   990
         TabIndex        =   34
         Top             =   3480
         Width           =   465
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   22
         Left            =   990
         TabIndex        =   33
         Top             =   3840
         Width           =   420
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Artículo"
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
         Left            =   600
         TabIndex        =   32
         Top             =   3240
         Width           =   540
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   7
         Left            =   1560
         ToolTipText     =   "Buscar artículo"
         Top             =   3480
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   8
         Left            =   1560
         ToolTipText     =   "Buscar artículo"
         Top             =   3870
         Width           =   240
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
         Index           =   3
         Left            =   600
         TabIndex        =   29
         Top             =   4320
         Width           =   660
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   6
         Left            =   1545
         MouseIcon       =   "frmEstTarArt.frx":000C
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar colectivo"
         Top             =   4320
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   5
         Left            =   1545
         MouseIcon       =   "frmEstTarArt.frx":015E
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar familia"
         Top             =   2895
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   4
         Left            =   1560
         MouseIcon       =   "frmEstTarArt.frx":02B0
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar familia"
         Top             =   2520
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Familia"
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
         TabIndex        =   22
         Top             =   2280
         Width           =   480
      End
      Begin VB.Label Label4 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   1
         Left            =   960
         TabIndex        =   21
         Top             =   2895
         Width           =   420
      End
      Begin VB.Label Label4 
         Caption         =   "Desde"
         Height          =   195
         Index           =   0
         Left            =   960
         TabIndex        =   20
         Top             =   2520
         Width           =   465
      End
      Begin VB.Label Label4 
         Caption         =   "Fecha Albarán"
         ForeColor       =   &H00972E0B&
         Height          =   255
         Index           =   16
         Left            =   600
         TabIndex        =   19
         Top             =   1320
         Width           =   1815
      End
      Begin VB.Label Label4 
         Caption         =   "Desde"
         Height          =   195
         Index           =   15
         Left            =   960
         TabIndex        =   18
         Top             =   1560
         Width           =   465
      End
      Begin VB.Label Label4 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   14
         Left            =   960
         TabIndex        =   17
         Top             =   1920
         Width           =   420
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   2
         Left            =   1530
         Picture         =   "frmEstTarArt.frx":0402
         ToolTipText     =   "Buscar fecha"
         Top             =   1560
         Width           =   240
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   3
         Left            =   1530
         Picture         =   "frmEstTarArt.frx":048D
         ToolTipText     =   "Buscar fecha"
         Top             =   1920
         Width           =   240
      End
      Begin VB.Label Label4 
         Caption         =   "Desde"
         Height          =   195
         Index           =   13
         Left            =   960
         TabIndex        =   16
         Top             =   600
         Width           =   465
      End
      Begin VB.Label Label4 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   12
         Left            =   960
         TabIndex        =   15
         Top             =   975
         Width           =   420
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Tarjeta"
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
         TabIndex        =   14
         Top             =   360
         Width           =   525
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   0
         Left            =   1560
         MouseIcon       =   "frmEstTarArt.frx":0518
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar tarjeta"
         Top             =   600
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   1
         Left            =   1545
         MouseIcon       =   "frmEstTarArt.frx":066A
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar tarjeta"
         Top             =   975
         Width           =   240
      End
   End
End
Attribute VB_Name = "frmEstTarArt"
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

Private WithEvents frmFam As frmManFamia 'Familias
Attribute frmFam.VB_VarHelpID = -1
Private WithEvents frmCol As frmManCoope 'Colectivos
Attribute frmCol.VB_VarHelpID = -1
Private WithEvents frmTar As frmManTarje 'Tarjetas
Attribute frmTar.VB_VarHelpID = -1
Private WithEvents frmC As frmCal 'calendario fechas
Attribute frmC.VB_VarHelpID = -1
Private WithEvents frmArt As frmManArtic 'Articulos
Attribute frmArt.VB_VarHelpID = -1

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

Private Sub cmdAceptar_Click()
Dim cDesde As String, cHasta As String 'cadena codigo Desde/Hasta
Dim nDesde As String, nHasta As String 'cadena Descripcion Desde/Hasta
Dim cadTabla As String, cOrden As String
Dim I As Byte
InicializarVbles
    
    If txtCodigo(6).Text = "" Then
        MsgBox "Introduzca el Colectivo a listar.", vbExclamation
        Exit Sub
    End If
    
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
        Codigo = "{" & tabla & ".numtarje}"
        TipCod = "N"
        If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHcliente= """) Then Exit Sub
    End If
    
    'D/H Fecha albaran
    cDesde = Trim(txtCodigo(2).Text)
    cHasta = Trim(txtCodigo(3).Text)
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{" & tabla & ".fecalbar}"
        TipCod = "F"
        If Not PonerDesdeHasta(cDesde, cHasta, "", "", "pDHfechaFac= """) Then Exit Sub
    End If
    
    'D/H Familia
    cDesde = Trim(txtCodigo(4).Text)
    cHasta = Trim(txtCodigo(5).Text)
    nDesde = txtNombre(4).Text
    nHasta = txtNombre(5).Text
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{sartic.codfamia}"
        TipCod = "N"
        If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHFami= """) Then Exit Sub
    End If
    
    '[Monica]29/08/2017: añadido el desde/hasta articulo
    'D/H Articulo
    cDesde = Trim(txtCodigo(7).Text)
    cHasta = Trim(txtCodigo(8).Text)
    nDesde = txtNombre(7).Text
    nHasta = txtNombre(8).Text
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{sartic.codartic}"
        TipCod = "N"
        If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHArtic= """) Then Exit Sub
    End If
    
    
    'Colectivo
    cDesde = Trim(txtCodigo(6).Text)
    cHasta = Trim(txtCodigo(6).Text)
    nDesde = txtNombre(6).Text
    nHasta = txtNombre(6).Text
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{schfac.codcoope}"
        TipCod = "N"
        If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHCole= """) Then Exit Sub
    End If
    
    'Comprobar si hay registros a Mostrar antes de abrir el Informe
    cadTabla = "((" & "(" & tabla & " INNER JOIN starje ON " & tabla & ".numtarje=starje.numtarje) "
    cadTabla = cadTabla & "inner join sartic on slhfac.codartic = sartic.codartic )"
    cadTabla = cadTabla & "inner join sfamia on sartic.codfamia = sfamia.codfamia )"
    cadTabla = cadTabla & "INNER JOIN schfac ON " & "slhfac.letraser=schfac.letraser AND slhfac.numfactu=schfac.numfactu AND slhfac.fecfactu=schfac.fecfactu"
    
    If HayRegParaInforme(cadTabla, cadSelect) Then
       If Option1(0) = True Then
          cadTitulo = "Ventas Artículos por Tarjeta"
          cadNombreRPT = "rEstTarart.rpt"
          LlamarImprimir
       Else
          cadTitulo = "Resumen Ventas por Tarjeta"
          cadNombreRPT = "rEstTarart1.rpt"
          LlamarImprimir
       End If
    End If
    
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    If PrimeraVez Then
        PrimeraVez = False
        PonerFoco txtCodigo(0)
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
Dim h As Integer, w As Integer
Dim List As Collection

    PrimeraVez = True
    limpiar Me

    'IMAGES para busqueda
     Me.imgBuscar(0).Picture = frmPpal.imgListImages16.ListImages(1).Picture
     Me.imgBuscar(1).Picture = frmPpal.imgListImages16.ListImages(1).Picture
     Me.imgBuscar(4).Picture = frmPpal.imgListImages16.ListImages(1).Picture
     Me.imgBuscar(5).Picture = frmPpal.imgListImages16.ListImages(1).Picture
     Me.imgBuscar(6).Picture = frmPpal.imgListImages16.ListImages(1).Picture
     Me.imgBuscar(7).Picture = frmPpal.imgListImages16.ListImages(1).Picture
     Me.imgBuscar(8).Picture = frmPpal.imgListImages16.ListImages(1).Picture

    '###Descomentar
'    CommitConexion
         
    FrameCobrosVisible True, h, w
    indFrame = 5
    tabla = "slhfac"
            
    'Esto se consigue poneinedo el cancel en el opcion k corresponda
    Me.cmdCancel.Cancel = True
End Sub

Private Sub frmArt_DatoSeleccionado(CadenaSeleccion As String)
'Form de Consulta de articulos
    txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000000")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmC_Selec(vFecha As Date)
 'Fecha
    txtCodigo(CByte(imgFec(2).Tag)).Text = Format(vFecha, "dd/MM/yyyy")
End Sub

Private Sub frmTar_DatoSeleccionado(CadenaSeleccion As String)
'Form de Consulta de Tarjetas
    txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "00000000")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmFam_DatoSeleccionado(CadenaSeleccion As String)
'Form de Consulta de Familias
    txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "00")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmCol_DatoSeleccionado(CadenaSeleccion As String)
'Form de Consulta de Colectivos
    txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "00")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
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

Private Sub imgBuscar_Click(Index As Integer)
   Select Case Index
        Case 0, 1 'TARJETAS
            AbrirfrmTarjetas (Index)
        
        Case 4, 5 'FAMILIAS
            AbrirFrmFamilias (Index)
        
        Case 6 'COLECTIVO
            AbrirFrmColectivos (Index)
        
        Case 7, 8 'ARTICULOS
            AbrirFrmArticulos (Index)
        
    End Select
    PonerFoco txtCodigo(indCodigo)
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
' ahora
    If KeyAscii = teclaBuscar Then
        Select Case Index
            Case 0: KEYBusqueda KeyAscii, 0 'tarjeta desde
            Case 1: KEYBusqueda KeyAscii, 1 'tarjeta hasta
            Case 6: KEYBusqueda KeyAscii, 6 'colectivo
            Case 4: KEYBusqueda KeyAscii, 4 'familia desde
            Case 5: KEYBusqueda KeyAscii, 5 'familia hasta
            Case 2: KEYFecha KeyAscii, 2 'fecha desde
            Case 3: KEYFecha KeyAscii, 3 'fecha hasta
            
            Case 7: KEYBusqueda KeyAscii, 7 'articulo desde
            Case 8: KEYBusqueda KeyAscii, 8 'articulo hasta
            
            
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
Dim Cad As String, cadTipo As String 'tipo cliente

    'Quitar espacios en blanco por los lados
    txtCodigo(Index).Text = Trim(txtCodigo(Index).Text)
    
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub

    
    Select Case Index
            
        Case 0, 1 'TARJETAS
            txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), "starje", "nomtarje", "numtarje", "N")
            If txtCodigo(Index).Text <> "" Then txtCodigo(Index).Text = Format(txtCodigo(Index).Text, "00000000")
        
        Case 2, 3 'FECHAS
            If txtCodigo(Index).Text <> "" Then PonerFormatoFecha txtCodigo(Index)
            
        Case 4, 5 'FAMILIAS
            txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), "sfamia", "nomfamia", "codfamia", "N")
            If txtCodigo(Index).Text <> "" Then txtCodigo(Index).Text = Format(txtCodigo(Index).Text, "00")
        
        Case 6 'COLECTIVOS
            txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), "scoope", "nomcoope", "codcoope", "N")
            If txtCodigo(Index).Text <> "" Then txtCodigo(Index).Text = Format(txtCodigo(Index).Text, "00")
            
        Case 7, 8 'ARTICULOS
            txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), "sartic", "nomartic", "codartic", "N")
            If txtCodigo(Index).Text <> "" Then txtCodigo(Index).Text = Format(txtCodigo(Index).Text, "000000")
            
    End Select
End Sub

Private Sub FrameCobrosVisible(visible As Boolean, ByRef h As Integer, ByRef w As Integer)
    Me.FrameCobros.visible = visible
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
        .Opcion = 0
        .Show vbModal
    End With
End Sub

Private Sub AbrirfrmTarjetas(indice As Integer)
    indCodigo = indice
    Set frmTar = New frmManTarje
    frmTar.DatosADevolverBusqueda = "3|4|"
    frmTar.DeConsulta = True
    frmTar.CodigoActual = DevuelveDesdeBDNew(cPTours, "starje", "codsocio", "numtarje", txtCodigo(indCodigo), "N")
    frmTar.Show vbModal
    Set frmTar = Nothing
End Sub

Private Sub AbrirFrmFamilias(indice As Integer)
    indCodigo = indice
    Set frmFam = New frmManFamia
    frmFam.DatosADevolverBusqueda = "0|1|"
    frmFam.DeConsulta = True
    frmFam.CodigoActual = txtCodigo(indCodigo)
    frmFam.Show vbModal
    Set frmFam = Nothing
End Sub

Private Sub AbrirFrmColectivos(indice As Integer)
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

Private Sub AbrirFrmArticulos(indice As Integer)
    indCodigo = indice
    Set frmArt = New frmManArtic
    frmArt.DatosADevolverBusqueda = "0|1|"
    frmArt.DeConsulta = True
    frmArt.CodigoActual = txtCodigo(indCodigo)
    frmArt.Show vbModal
    Set frmArt = Nothing
End Sub

