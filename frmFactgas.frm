VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmFactgas 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Reimpresion de Facturas"
   ClientHeight    =   6495
   ClientLeft      =   45
   ClientTop       =   2430
   ClientWidth     =   7185
   Icon            =   "frmFactgas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6495
   ScaleWidth      =   7185
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
      Width           =   6915
      Begin VB.CheckBox ChkTipoDocu 
         Caption         =   "Incluir las que se envian por email"
         Height          =   255
         Index           =   2
         Left            =   780
         TabIndex        =   33
         Top             =   5430
         Value           =   1  'Checked
         Width           =   2925
      End
      Begin VB.CheckBox ChkTipoDocu 
         Caption         =   "Facturas CEPSA"
         Height          =   255
         Index           =   1
         Left            =   3780
         TabIndex        =   32
         Top             =   4830
         Value           =   1  'Checked
         Width           =   2355
      End
      Begin VB.CheckBox ChkTipoDocu 
         Caption         =   "Facturas con Gasóleo B"
         Height          =   255
         Index           =   0
         Left            =   3780
         TabIndex        =   31
         Top             =   4470
         Value           =   1  'Checked
         Width           =   2355
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   8
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   29
         Text            =   "Text5"
         Top             =   600
         Width           =   3135
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   8
         Left            =   1845
         MaxLength       =   6
         TabIndex        =   0
         Top             =   600
         Width           =   830
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   7
         Left            =   4245
         MaxLength       =   7
         TabIndex        =   8
         Top             =   3720
         Width           =   930
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   6
         Left            =   4245
         MaxLength       =   7
         TabIndex        =   7
         Top             =   3360
         Width           =   930
      End
      Begin VB.Frame Frame4 
         Caption         =   "Facturas"
         ForeColor       =   &H00972E0B&
         Height          =   1000
         Left            =   900
         TabIndex        =   25
         Top             =   4260
         Width           =   2175
         Begin VB.OptionButton Option1 
            Caption         =   "Clientes marcados"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   9
            Top             =   240
            Value           =   -1  'True
            Width           =   1815
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Todos"
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   10
            Top             =   600
            Width           =   1815
         End
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   5
         Left            =   1845
         MaxLength       =   3
         TabIndex        =   6
         Top             =   3735
         Width           =   570
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   4
         Left            =   1845
         MaxLength       =   3
         TabIndex        =   5
         Top             =   3360
         Width           =   570
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   3
         Left            =   1845
         MaxLength       =   10
         TabIndex        =   4
         Tag             =   "Código Postal|T|S|||clientes|codposta|||"
         Top             =   2640
         Width           =   1050
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   2
         Left            =   1845
         MaxLength       =   10
         TabIndex        =   3
         Tag             =   "Código Postal|T|S|||clientes|codposta|||"
         Top             =   2310
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
         Left            =   1845
         MaxLength       =   6
         TabIndex        =   1
         Top             =   1200
         Width           =   830
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   1
         Left            =   1845
         MaxLength       =   6
         TabIndex        =   2
         Top             =   1590
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
         Top             =   1200
         Width           =   3135
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
         Top             =   1575
         Width           =   3135
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   8
         Left            =   1560
         MouseIcon       =   "frmFactgas.frx":000C
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar cliente"
         Top             =   600
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
         Index           =   6
         Left            =   600
         TabIndex        =   30
         Top             =   360
         Width           =   660
      End
      Begin VB.Label Label4 
         Caption         =   "Desde"
         Height          =   195
         Index           =   5
         Left            =   3360
         TabIndex        =   28
         Top             =   3360
         Width           =   465
      End
      Begin VB.Label Label4 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   4
         Left            =   3360
         TabIndex        =   27
         Top             =   3735
         Width           =   420
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Factura"
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
         Left            =   3000
         TabIndex        =   26
         Top             =   3120
         Width           =   555
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Serie"
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
         Top             =   3120
         Width           =   360
      End
      Begin VB.Label Label4 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   1
         Left            =   960
         TabIndex        =   21
         Top             =   3735
         Width           =   420
      End
      Begin VB.Label Label4 
         Caption         =   "Desde"
         Height          =   195
         Index           =   0
         Left            =   960
         TabIndex        =   20
         Top             =   3360
         Width           =   465
      End
      Begin VB.Label Label4 
         Caption         =   "Fecha Factura"
         ForeColor       =   &H00972E0B&
         Height          =   255
         Index           =   16
         Left            =   600
         TabIndex        =   19
         Top             =   2040
         Width           =   1815
      End
      Begin VB.Label Label4 
         Caption         =   "Desde"
         Height          =   195
         Index           =   15
         Left            =   960
         TabIndex        =   18
         Top             =   2280
         Width           =   465
      End
      Begin VB.Label Label4 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   14
         Left            =   960
         TabIndex        =   17
         Top             =   2640
         Width           =   420
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   2
         Left            =   1530
         Picture         =   "frmFactgas.frx":015E
         ToolTipText     =   "Buscar fecha"
         Top             =   2280
         Width           =   240
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   3
         Left            =   1530
         Picture         =   "frmFactgas.frx":01E9
         ToolTipText     =   "Buscar fecha"
         Top             =   2640
         Width           =   240
      End
      Begin VB.Label Label4 
         Caption         =   "Desde"
         Height          =   195
         Index           =   13
         Left            =   960
         TabIndex        =   16
         Top             =   1200
         Width           =   465
      End
      Begin VB.Label Label4 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   12
         Left            =   960
         TabIndex        =   15
         Top             =   1575
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
         TabIndex        =   14
         Top             =   960
         Width           =   495
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   0
         Left            =   1560
         MouseIcon       =   "frmFactgas.frx":0274
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar cliente"
         Top             =   1200
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   1
         Left            =   1545
         MouseIcon       =   "frmFactgas.frx":03C6
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar cliente"
         Top             =   1575
         Width           =   240
      End
   End
End
Attribute VB_Name = "frmFactgas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' +-+-+-+-+-+-+-+-+-+-+-
' +-+- Autor MANOLO +-+-
' +-+-+-+-+-+-+-+-+-+-+-

Option Explicit

Public NumCod As String ' 0 factura normal
                        ' 1 ajena
                        

Public CadTag As String 'Cadena con el Tag del campo que se va a poner en D/H en los listados
                        'Se necesita si el tipo de codigo es texto

    
Private Conexion As Byte
'1.- Conexión a BD Ariges  2.- Conexión a BD Conta

Private HaDevueltoDatos As Boolean

Private WithEvents frmcli As frmManClien 'Clientes
Attribute frmcli.VB_VarHelpID = -1
Private WithEvents frmCol As frmManCoope 'Colectivo
Attribute frmCol.VB_VarHelpID = -1
Private WithEvents frmC As frmCal 'calendario fechas
Attribute frmC.VB_VarHelpID = -1

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

Dim OpcionListado As Byte


Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, 0, cerrar
    If cerrar Then Unload Me
End Sub

Private Sub cmdAceptar_Click()
Dim cDesde As String, cHasta As String 'cadena codigo Desde/Hasta
Dim nDesde As String, nHasta As String 'cadena Descripcion Desde/Hasta
Dim cadTABLA As String, cOrden As String
Dim i As Byte
Dim b As Boolean
Dim Letra As String

InicializarVbles
    
'    If txtCodigo(6).Text = "" Then
'        MsgBox "Introduzca el Colectivo a listar.", vbExclamation
'        Exit Sub
'    End If
    
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
        Codigo = "{" & tabla & ".fecfactu}"
        TipCod = "F"
        If Not PonerDesdeHasta(cDesde, cHasta, "", "", "pDHfechaFac= """) Then Exit Sub
    End If
    
    'D/H Serie
    cDesde = Trim(txtCodigo(4).Text)
    cHasta = Trim(txtCodigo(5).Text)
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{" & tabla & ".letraser}"
        TipCod = "T"
        If Not PonerDesdeHasta(cDesde, cHasta, "", "", "pDHSerie= """) Then Exit Sub
    End If
    
    'Factura
    cDesde = Trim(txtCodigo(6).Text)
    cHasta = Trim(txtCodigo(7).Text)
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{" & tabla & ".numfactu}"
        TipCod = "N"
        If Not PonerDesdeHasta(cDesde, cHasta, "", "", "pDHFact= """) Then Exit Sub
    End If
    
    
    If NumCod = 1 Then
        b = AnyadirAFormula(cadFormula, "{" & tabla & ".codcoope} = " & txtCodigo(8).Text)
        b = AnyadirAFormula(cadSelect, tabla & ".codcoope = " & DBSet(txtCodigo(8).Text, "N"))
    End If
    
    
    'Tipo de Socio
    Codigo = "{ssocio.impfactu}"
    cDesde = ""
    cHasta = ""
    If Option1(0) = True Then
        cDesde = Codigo & "=1"
    Else
        cDesde = Codigo & "<=1"
    End If
    
    cDesde = "(" & cDesde & ")"
    AnyadirAFormula cadFormula, cDesde
    AnyadirAFormula cadSelect, cDesde
    'Añadir el parametro tipo documentos seleccionados
    cadParam = cadParam & "pTipoDoc=""" & cHasta & """|"
    numParam = numParam + 1
    
    'Comprobar si hay registros a Mostrar antes de abrir el Informe
    cadTABLA = tabla & " INNER JOIN ssocio ON " & tabla & ".codsocio=ssocio.codsocio "
    If ChkTipoDocu(0).Value = 1 Then  ' unicamente seleccionamos las facturas que tengan algun articulo de gasoleo B
        If NumCod = 0 Then
            cadTABLA = "((" & Trim(cadTABLA) & ")"
            cadTABLA = cadTABLA & " inner join slhfac on schfac.letraser = slhfac.letraser and schfac.numfactu = slhfac.numfactu "
            cadTABLA = cadTABLA & " and schfac.fecfactu = slhfac.fecfactu) inner join sartic on slhfac.codartic = sartic.codartic and sartic.tipogaso = 3 "
        Else
            cadTABLA = "((" & Trim(cadTABLA) & ")"
            cadTABLA = cadTABLA & " inner join slhfacr on schfacr.letraser = slhfacr.letraser and schfacr.numfactu = slhfacr.numfactu "
            cadTABLA = cadTABLA & " and schfacr.fecfactu = slhfacr.fecfactu) inner join sartic on slhfacr.codartic = sartic.codartic and sartic.tipogaso = 3 "
        End If
        AnyadirAFormula cadFormula, "{tmpfacturas.codusu}= " & vSesion.Codigo
    End If

    '[Monica]15/12/2010: solo los que tengan la letra de serie de FAC ( factura cepsa ) en Pobla del duc
    If vParamAplic.Cooperativa = 4 Then
        Letra = DevuelveDesdeBDNew(cPTours, "stipom", "letraser", "codtipom", "FAC", "T")
        If ChkTipoDocu(1).Value = 1 Then
            cDesde = "({schfac.letraser} = '" & Letra & "')"
            AnyadirAFormula cadFormula, cDesde
            AnyadirAFormula cadSelect, cDesde
        Else
            cDesde = "({schfac.letraser} <> '" & Letra & "')"
            AnyadirAFormula cadFormula, cDesde
            AnyadirAFormula cadSelect, cDesde
        End If
    End If

    '[Monica]28/12/2011: incluir la impresion de facturas de los socios que tienen la impresion por email
    '                    si está marcado no hago nada
    '                    si no está marcado quito la impresion de facturas que salen por email
    If ChkTipoDocu(2).Value = 0 Then
        If NumCod = 0 Then
            cadTABLA = "(" & Trim(cadTABLA) & ")"
            cadTABLA = cadTABLA & " inner join ssocio on schfac.codsocio = ssocio.codsocio and ssocio.envfactemail = 0 "
        Else
            cadTABLA = "((" & Trim(cadTABLA) & ")"
            cadTABLA = cadTABLA & " inner join ssocio on schfacr.codsocio = ssocio.codsocio and ssocio.envfactemail = 0 "
        End If
    End If


    If HayRegParaInforme(cadTABLA, cadSelect) Then
        If Not InsertarNrosFacturaEnTemporal(cadTABLA, cadSelect) Then Exit Sub
    
        '23022007 Monica solo en el caso de Alzira quieren las bonificaciones después de las lineas
        '[Monica]09/01/2013: Nueva cooperativa de Ribarroja
        If vParamAplic.Cooperativa = 1 Or vParamAplic.Cooperativa = 5 Then cadFormula = cadFormula & " and {slhfac.numalbar} <> ""BONIFICA"""
        
        If NumCod = 0 Then
            cadTitulo = "Reimpresion de Facturas"
        Else
            cadTitulo = "Reimpresion de Facturas Ajenas"
        End If
       
        ' ### [Monica] 11/09/2006
        '****************************
        Dim indRPT As Byte 'Indica el tipo de Documento en la tabla "scryst"
        Dim nomDocu As String 'Nombre de Informe rpt de crystal
        
        indRPT = 1 'Facturas Clientes
        
       If Not PonerParamRPT(indRPT, cadParam, numParam, nomDocu) Then Exit Sub
       'Nombre fichero .rpt a Imprimir
       If NumCod = 1 Then
            nomDocu = Replace(nomDocu, ".rpt", "Aj" & "C" & Format(txtCodigo(8).Text, "00") & ".rpt")
       End If
       
       If ChkTipoDocu(1).Value = 1 Then nomDocu = Replace(nomDocu, ".rpt", "Cepsa.rpt")
       
       
       ' he añadido estas dos lineas para que llame al rpt correspondiente
       frmImprimir.NombreRPT = nomDocu
       cadNombreRPT = nomDocu  ' "rFactgas.rpt"
       '[Monica]09/01/2013: Nueva cooperativa de Ribarroja
       If vParamAplic.Cooperativa = 1 Or vParamAplic.Cooperativa = 5 Then
            OpcionListado = 1
       Else
            OpcionListado = 0
       End If
       LlamarImprimir
    End If
    
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    If PrimeraVez Then
        PrimeraVez = False
        
        If NumCod = 1 Then
            PonerFoco txtCodigo(8)
            Option1(1).Value = True
        Else
            PonerFoco txtCodigo(0)
        End If
        
      ' PonerFoco txtCodigo(0)
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
Dim h As Integer, w As Integer
Dim List As Collection

    PrimeraVez = True
    limpiar Me

    Label4(6).visible = (NumCod = 1)
    txtCodigo(8).visible = (NumCod = 1)
    txtCodigo(8).Enabled = (NumCod = 1)
    txtNombre(8).visible = (NumCod = 1)
    txtNombre(8).Enabled = (NumCod = 1)
    Me.imgBuscar(8).visible = (NumCod = 1)
    Me.imgBuscar(8).Enabled = (NumCod = 1)
    ChkTipoDocu(0).Value = 0
    
    ChkTipoDocu(1).visible = (vParamAplic.Cooperativa = 4)
    ChkTipoDocu(1).Enabled = (vParamAplic.Cooperativa = 4)
    ChkTipoDocu(1).Value = 0

    'IMAGES para busqueda
     Me.imgBuscar(0).Picture = frmPpal.imgListImages16.ListImages(1).Picture
     Me.imgBuscar(1).Picture = frmPpal.imgListImages16.ListImages(1).Picture
     Me.imgBuscar(8).Picture = frmPpal.imgListImages16.ListImages(1).Picture

    '###Descomentar
'    CommitConexion
         
    FrameCobrosVisible True, h, w
    indFrame = 5
    If NumCod = 0 Then
        tabla = "schfac"
    Else
        Me.Caption = Me.Caption & " Ajenas"
        tabla = "schfacr" ' historico del Regaixo
    End If
            
    'Esto se consigue poneinedo el cancel en el opcion k corresponda
    Me.cmdCancel.Cancel = True



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
        Case 0, 1 'CLIENTE
            AbrirFrmClientes (Index)
        
        Case 8 ' cooperativa
            AbrirFrmColectivo (Index)
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
'ahora
    If KeyAscii = teclaBuscar Then
        Select Case Index
            Case 0: KEYBusqueda KeyAscii, 0 'cliente desde
            Case 1: KEYBusqueda KeyAscii, 1 'cliente hasta
            Case 6: KEYBusqueda KeyAscii, 6 'numero de factura desde
            Case 7: KEYBusqueda KeyAscii, 7 'numero de factura hasta
            Case 2: KEYFecha KeyAscii, 3 'fecha desde
            Case 3: KEYFecha KeyAscii, 3 'fecha hasta
            Case 8: KEYBusqueda KeyAscii, 8 'colectivo
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
            
        Case 4, 5 'SERIE
            txtCodigo(Index).Text = UCase(txtCodigo(Index).Text)
        
        Case 6, 7 'FACTURAS
            txtCodigo(Index).Text = Format(txtCodigo(Index).Text, "0000000")
            
        Case 8 'COLECTIVO
            txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), "scoope", "nomcoope", "codcoope", "N")
            If txtCodigo(Index).Text <> "" Then txtCodigo(Index).Text = Format(txtCodigo(Index).Text, "000")
        
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
        .Opcion = OpcionListado
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
'       .ExportarPDF = (chkEMAIL.Value = 1)
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

Public Function InsertarNrosFacturaEnTemporal(cTabla As String, cWhere As String) As Boolean
'Comprobar si hay registros a Mostrar antes de abrir el Informe
Dim sql As String
Dim sql2 As String

    On Error GoTo eInsertarNrosFacturaEnTemporal

    InsertarNrosFacturaEnTemporal = False

    sql2 = "delete from tmpfacturas where codusu = " & vSesion.Codigo
    Conn.Execute sql2

    sql = "Select distinct " & vSesion.Codigo & "," & tabla & ".letraser," & tabla & ".numfactu" & "," & tabla & ".fecfactu " & " FROM " & QuitarCaracterACadena(cTabla, "_1")
    If cWhere <> "" Then
        cWhere = QuitarCaracterACadena(cWhere, "{")
        cWhere = QuitarCaracterACadena(cWhere, "}")
        cWhere = QuitarCaracterACadena(cWhere, "_1")
        sql = sql & " WHERE " & cWhere
    End If
    
    sql2 = "insert into tmpfacturas (codusu, letraser, numfactu, fecfactu) " & sql
    Conn.Execute sql2
    
eInsertarNrosFacturaEnTemporal:
    If Err.Number = 0 Then
        InsertarNrosFacturaEnTemporal = True
    Else
        MsgBox "Se ha producido un error en la carga de la tabla intermedia", vbExclamation
    End If
End Function

