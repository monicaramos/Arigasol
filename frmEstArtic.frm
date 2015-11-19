VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmEstArtic 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Resumen Ventas Artículos"
   ClientHeight    =   5625
   ClientLeft      =   45
   ClientTop       =   2430
   ClientWidth     =   7185
   Icon            =   "frmEstArtic.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5625
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
      Height          =   5460
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   6915
      Begin VB.Frame Frame4 
         Caption         =   "Tipo Informe"
         ForeColor       =   &H00972E0B&
         Height          =   885
         Left            =   495
         TabIndex        =   19
         Top             =   3630
         Width           =   2235
         Begin VB.OptionButton Option1 
            Caption         =   "Detalle Colectivo"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   21
            Top             =   240
            Value           =   -1  'True
            Width           =   1815
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Resumen"
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   20
            Top             =   540
            Width           =   1815
         End
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   4
         Left            =   1875
         MaxLength       =   6
         TabIndex        =   4
         Top             =   2850
         Width           =   830
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   5
         Left            =   1875
         MaxLength       =   6
         TabIndex        =   5
         Top             =   3210
         Width           =   830
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   4
         Left            =   2775
         Locked          =   -1  'True
         TabIndex        =   18
         Text            =   "Text5"
         Top             =   2835
         Width           =   3135
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   5
         Left            =   2775
         Locked          =   -1  'True
         TabIndex        =   17
         Text            =   "Text5"
         Top             =   3225
         Width           =   3135
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   3
         Left            =   1845
         MaxLength       =   10
         TabIndex        =   3
         Tag             =   "Código Postal|T|S|||clientes|codposta|||"
         Top             =   2280
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
         Top             =   1920
         Width           =   1050
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   5235
         TabIndex        =   7
         Top             =   4905
         Width           =   975
      End
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   4050
         TabIndex        =   6
         Top             =   4905
         Width           =   975
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   0
         Left            =   1845
         MaxLength       =   2
         TabIndex        =   0
         Top             =   840
         Width           =   830
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   1
         Left            =   1845
         MaxLength       =   2
         TabIndex        =   1
         Top             =   1215
         Width           =   830
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   2745
         Locked          =   -1  'True
         TabIndex        =   10
         Text            =   "Text5"
         Top             =   840
         Width           =   3135
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   2745
         Locked          =   -1  'True
         TabIndex        =   9
         Text            =   "Text5"
         Top             =   1215
         Width           =   3135
      End
      Begin MSComctlLib.ProgressBar Pb1 
         Height          =   255
         Left            =   510
         TabIndex        =   25
         Top             =   4590
         Width           =   5670
         _ExtentX        =   10001
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Label Label4 
         Caption         =   "Cargando temporal"
         Height          =   195
         Index           =   3
         Left            =   570
         TabIndex        =   26
         Top             =   4890
         Width           =   2235
      End
      Begin VB.Label Label4 
         Caption         =   "Desde"
         Height          =   195
         Index           =   0
         Left            =   990
         TabIndex        =   24
         Top             =   2850
         Width           =   465
      End
      Begin VB.Label Label4 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   1
         Left            =   990
         TabIndex        =   23
         Top             =   3225
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
         Index           =   2
         Left            =   630
         TabIndex        =   22
         Top             =   2610
         Width           =   660
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   4
         Left            =   1575
         MouseIcon       =   "frmEstArtic.frx":000C
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar colectivo"
         Top             =   2850
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   5
         Left            =   1575
         MouseIcon       =   "frmEstArtic.frx":015E
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar colectivo"
         Top             =   3225
         Width           =   240
      End
      Begin VB.Label Label4 
         Caption         =   "Fecha factura"
         ForeColor       =   &H00972E0B&
         Height          =   255
         Index           =   16
         Left            =   600
         TabIndex        =   16
         Top             =   1680
         Width           =   1815
      End
      Begin VB.Label Label4 
         Caption         =   "Desde"
         Height          =   195
         Index           =   15
         Left            =   960
         TabIndex        =   15
         Top             =   1920
         Width           =   465
      End
      Begin VB.Label Label4 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   14
         Left            =   960
         TabIndex        =   14
         Top             =   2280
         Width           =   420
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   2
         Left            =   1530
         Picture         =   "frmEstArtic.frx":02B0
         ToolTipText     =   "Buscar fecha"
         Top             =   1920
         Width           =   240
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   3
         Left            =   1530
         Picture         =   "frmEstArtic.frx":033B
         ToolTipText     =   "Buscar fecha"
         Top             =   2280
         Width           =   240
      End
      Begin VB.Label Label4 
         Caption         =   "Desde"
         Height          =   195
         Index           =   13
         Left            =   960
         TabIndex        =   13
         Top             =   840
         Width           =   465
      End
      Begin VB.Label Label4 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   12
         Left            =   960
         TabIndex        =   12
         Top             =   1215
         Width           =   420
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
         Index           =   11
         Left            =   600
         TabIndex        =   11
         Top             =   600
         Width           =   480
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   0
         Left            =   1560
         MouseIcon       =   "frmEstArtic.frx":03C6
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar familia"
         Top             =   840
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   1
         Left            =   1575
         MouseIcon       =   "frmEstArtic.frx":0518
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar familia"
         Top             =   1215
         Width           =   240
      End
   End
End
Attribute VB_Name = "frmEstArtic"
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

Private Sub cmdAceptar_Click()
Dim cDesde As String, cHasta As String 'cadena codigo Desde/Hasta
Dim nDesde As String, nHasta As String 'cadena Descripcion Desde/Hasta
Dim cadTABLA As String, cOrden As String
Dim i As Byte
InicializarVbles
    
    '========= PARAMETROS  =============================
    'Añadir el parametro de Empresa
    cadParam = cadParam & "|pEmpresa=""" & vEmpresa.nomEmpre & """|"
    numParam = numParam + 1
    
    'D/H Familia
    cDesde = Trim(txtCodigo(0).Text)
    cHasta = Trim(txtCodigo(1).Text)
    nDesde = txtNombre(0).Text
    nHasta = txtNombre(1).Text
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{sartic.codfamia}"
        TipCod = "N"
        If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHFami= """) Then Exit Sub
    End If
    
    'D/H Colectivo
    cDesde = Trim(txtCodigo(4).Text)
    cHasta = Trim(txtCodigo(5).Text)
    nDesde = txtNombre(4).Text
    nHasta = txtNombre(5).Text
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{schfac.codcoope}"
        TipCod = "N"
        If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHColec= """) Then Exit Sub
    End If
    
    'D/H Fecha factura
    cDesde = Trim(txtCodigo(2).Text)
    cHasta = Trim(txtCodigo(3).Text)
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{slhfac.fecfactu}"
        TipCod = "F"
        If Not PonerDesdeHasta(cDesde, cHasta, "", "", "pDHfechaFac= """) Then Exit Sub
    End If
    
        
    'Comprobar si hay registros a Mostrar antes de abrir el Informe
    cadTABLA = "(" & tabla & " INNER JOIN schfac ON schfac.letraser = " & tabla & ".letraser and schfac.numfactu = " & tabla & ".numfactu and schfac.fecfactu = " & tabla & ".fecfactu) "
    cadTABLA = "(" & cadTABLA & " INNER JOIN sartic ON " & tabla & ".codartic=sartic.codartic) "
    cadTABLA = "(" & cadTABLA & " INNER JOIN sfamia ON " & "sartic.codfamia=sfamia.codfamia) "
    
'añadida la opcion
'    If Option1(0) = True Then
        If HayRegParaInforme(cadTABLA, cadSelect) Then
           If Option1(0) = True Then
              cadTitulo = "Resumen Ventas Colectivo Artículos"
              cadNombreRPT = "rEstArticDet.rpt"
              CargarTablaTemporal cDesde, cHasta, txtCodigo(0).Text, txtCodigo(1).Text, txtCodigo(4).Text, txtCodigo(5).Text
              cadFormula = "{tmpinformes.codusu} = " & vSesion.Codigo
              LlamarImprimir
              'AbrirVisReport
           Else
              cadTitulo = "Resumen Ventas Artículos"
              cadNombreRPT = "rEstArtic.rpt"
              CargarTablaTemporal cDesde, cHasta, txtCodigo(0).Text, txtCodigo(1).Text, txtCodigo(4).Text, txtCodigo(5).Text
              cadFormula = "{tmpinformes.codusu} = " & vSesion.Codigo
              LlamarImprimir
           End If
        End If
'    Else
'        If HayRegParaInforme(cadTABLA, cadSelect) Then
'            cadTitulo = "Resumen Ventas Artículos"
'            cadNombreRPT = "rEstArtic.rpt"
'            CargarTablaTemporal cDesde, cHasta, txtCodigo(0).Text, txtCodigo(1).Text
'            cadFormula = "{tmpinformes.codusu} = " & vSesion.Codigo
'            LlamarImprimir
'        End If
'    End If
'end añadido
    
'    If HayRegParaInforme(cadTABLA, cadSelect) Then
'       cadTitulo = "Resumen Ventas Artículos"
'       cadNombreRPT = "rEstArtic.rpt"
'       CargarTablaTemporal cDesde, cHasta, txtCodigo(0).Text, txtCodigo(1).Text
'       cadFormula = "{tmpinformes.codusu} = " & vSesion.Codigo
'       LlamarImprimir
'       'AbrirVisReport
'    End If
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

    '###Descomentar
'    CommitConexion
         
    FrameCobrosVisible True, h, w
    indFrame = 5
    tabla = "slhfac"
            
    'Esto se consigue poneinedo el cancel en el opcion k corresponda
    Me.cmdCancel.Cancel = True
'    Me.Width = w + 70
'    Me.Height = h + 350

    Me.Pb1.visible = False
    Me.Label4(3).visible = False
    

End Sub

Private Sub frmC_Selec(vFecha As Date)
 'Fecha
    txtCodigo(CByte(imgFec(2).Tag)).Text = Format(vFecha, "dd/MM/yyyy")
End Sub

Private Sub frmCol_DatoSeleccionado(CadenaSeleccion As String)
'Form de Consulta de Colectivos
    txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "00")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmFam_DatoSeleccionado(CadenaSeleccion As String)
'Form de Consulta de Familias
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

Private Sub ListView1_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub imgBuscar_Click(Index As Integer)
   Select Case Index
        
        Case 0, 1 'FAMILIAS
            AbrirFrmFamilias (Index)
        
        Case 4, 5 ' COLECTIVOS
            AbrirFrmColectivo (Index)
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

Private Sub txtCodigo_GotFocus(Index As Integer)
    ConseguirFoco txtCodigo(Index), 3
End Sub

Private Sub txtCodigo_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    KEYdown KeyCode
End Sub

Private Sub txtCodigo_KeyPress(Index As Integer, KeyAscii As Integer)
'14/02/2007 antes
'    KEYpress KeyAscii
' ahora
    If KeyAscii = teclaBuscar Then
        Select Case Index
            Case 0: KEYBusqueda KeyAscii, 0 'familia desde
            Case 1: KEYBusqueda KeyAscii, 1 'familia hasta
            Case 2: KEYFecha KeyAscii, 2 'fecha factura desde
            Case 3: KEYFecha KeyAscii, 3 'fecha factura hasta
            Case 4: KEYBusqueda KeyAscii, 4 'colectivo desde
            Case 5: KEYBusqueda KeyAscii, 5 'colectivo hasta
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
            
        Case 0, 1 'FAMILIAS
            txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), "sfamia", "nomfamia", "codfamia", "N")
            If txtCodigo(Index).Text <> "" Then txtCodigo(Index).Text = Format(txtCodigo(Index).Text, "00")
        
        Case 2, 3 'FECHAS
            If txtCodigo(Index).Text <> "" Then PonerFormatoFecha txtCodigo(Index)
        
        Case 4, 5 'COLECTIVO
            txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), "scoope", "nomcoope", "codcoope", "N")
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
        .Opcion = 0
        .Show vbModal
    End With
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
        '.SoloImprimir = (Me.OptVisualizar(indFrame).Value = 1)
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

Private Sub CargarTablaTemporal(DesFec As String, HasFec As String, DesFam As String, HasFam As String, DesCol As String, HasCol As String)
Dim SQL As String
Dim SQL1 As String
Dim Sql2 As String
Dim RS As ADODB.Recordset
Dim Importe As Currency
Dim vPorcIva As Currency
Dim ImporteSinIva As Currency
Dim vImpuesto As Currency
Dim SqlExiste As String
Dim CodIVA As String

    On Error GoTo eCargarTablaTemporal


    ' primero borramos los registros del usuario
    SQL = "delete from tmpinformes where codusu = " & DBSet(vSesion.Codigo, "N")
    Conn.Execute SQL

    '[Monica]05/02/2013: Añadimos la fecha de factura para saber que iva se le aplicó
    ' cargamos la tabla temporal para el listado agrupando por fecha y turno
    SQL = "select sartic.codfamia, schfac.codcoope, slhfac.codartic, sartic.codigiva, sartic.impuesto, sum(cantidad), sum(implinea), slhfac.fecfactu from slhfac, sartic, schfac where 1 = 1"
    If DesFec <> "" Then
        SQL = SQL & " and slhfac.fecfactu >= " & DBSet(DesFec, "F")
    End If
    If HasFec <> "" Then
        SQL = SQL & " and slhfac.fecfactu <= " & DBSet(HasFec, "F")
    End If
    If DesFam <> "" Then
        SQL = SQL & " and sartic.codfamia >= " & DBSet(DesFam, "N")
    End If
    If HasFam <> "" Then
        SQL = SQL & " and sartic.codfamia <= " & DBSet(HasFam, "N")
    End If
    If DesCol <> "" Then
        SQL = SQL & " and schfac.codcoope >= " & DBSet(DesCol, "N")
    End If
    If HasCol <> "" Then
        SQL = SQL & " and schfac.codcoope <= " & DBSet(HasCol, "N")
    End If
    SQL = SQL & " and slhfac.codartic = sartic.codartic "
    SQL = SQL & " and slhfac.letraser = schfac.letraser "
    SQL = SQL & " and slhfac.numfactu = schfac.numfactu "
    SQL = SQL & " and slhfac.fecfactu = schfac.fecfactu "
    SQL = SQL & " group by sartic.codfamia, schfac.codcoope, slhfac.codartic, codigiva, sartic.impuesto, slhfac.fecfactu "
    SQL = SQL & " order by sartic.codfamia, schfac.codcoope, slhfac.codartic, codigiva, sartic.impuesto, slhfac.fecfactu "
    
    Set RS = New ADODB.Recordset ' Crear objeto
    RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText ' abrir cursor
      
    Me.Pb1.visible = True
    Me.Label4(3).visible = True
    CargarProgres Pb1, TotalRegistrosConsulta(SQL)
    DoEvents
      
    If Not RS.EOF Then RS.MoveFirst
    
    While Not RS.EOF
        IncrementarProgres Pb1, 1
        DoEvents
        
        SqlExiste = "select count(*) from tmpinformes where codusu = " & vSesion.Codigo
        SqlExiste = SqlExiste & " and campo1 = " & DBSet(RS.Fields(0).Value, "N") ' familia
        SqlExiste = SqlExiste & " and campo2 = " & DBSet(RS.Fields(1).Value, "N") ' colectivo
        SqlExiste = SqlExiste & " and codigo1 = " & DBSet(RS.Fields(2).Value, "N") ' codartic
        
        If TotalRegistros(SqlExiste) <> 0 Then
            If DBLet(RS.Fields(7).Value, "F") < CDate(vParamAplic.FechaCamIva) Then
                If DBLet(RS.Fields(3).Value, "N") = vParamAplic.CodIvaGnral Then
                    CodIVA = vParamAplic.CodIvaGnralAnt
                ElseIf DBLet(RS.Fields(3).Value, "N") = vParamAplic.CodIvaRedu Then
                    CodIVA = vParamAplic.CodIvaReduAnt
                Else
                    CodIVA = DBLet(RS.Fields(3).Value, "N")
                End If
            Else
                CodIVA = DBLet(RS.Fields(3).Value, "N")
            End If
                
            Sql2 = ""
            Sql2 = DevuelveDesdeBDNew(cConta, "tiposiva", "porceiva", "codigiva", CodIVA, "N")
            If Sql2 = "" Then
                vPorcIva = 0
            Else
                vPorcIva = CCur(Sql2)
            End If
            
            ImporteSinIva = Round2(RS.Fields(6).Value / (1 + (vPorcIva / 100)), 2)
            If vParamAplic.Cooperativa = 2 Then ' si la cooperativa es el regaixo le quitamos el impuesto
                vImpuesto = Round2(DBLet(RS.Fields(5).Value, "N") * DBLet(RS.Fields(4).Value, "N"), 2)
                ImporteSinIva = ImporteSinIva - vImpuesto
            End If
            
            SQL1 = "update tmpinformes set importe1 = importe1 + (" & DBSet(RS.Fields(5).Value, "N") & ")"
            SQL1 = SQL1 & ", importe2 = importe2 + (" & DBSet(RS.Fields(6).Value, "N") & ")"
            SQL1 = SQL1 & ", importe3 = importe3 + (" & DBSet(ImporteSinIva, "N") & ")"
            SQL1 = SQL1 & " where codusu = " & DBSet(vSesion.Codigo, "N")
            SQL1 = SQL1 & " and campo1 = " & DBSet(RS.Fields(0).Value, "N")
            SQL1 = SQL1 & " and campo2 = " & DBSet(RS.Fields(1).Value, "N")
            SQL1 = SQL1 & " and codigo1 = " & DBSet(RS.Fields(2).Value, "N")
            
            Conn.Execute SQL1
        
        Else
            SQL1 = "insert into tmpinformes (codusu, campo1,  campo2, codigo1, importe1, importe2, importe3) "
            SQL1 = SQL1 & "values (" & DBSet(vSesion.Codigo, "N") & "," & DBSet(RS.Fields(0).Value, "N") & ","
            SQL1 = SQL1 & DBSet(RS.Fields(1).Value, "N") & "," ' colectivo
            SQL1 = SQL1 & DBSet(RS.Fields(2).Value, "N") & "," ' codartic
            SQL1 = SQL1 & DBSet(RS.Fields(5).Value, "N") & "," ' cantidad
            SQL1 = SQL1 & DBSet(RS.Fields(6).Value, "N") & "," ' importe con iva
            
            If DBLet(RS.Fields(7).Value, "F") < CDate(vParamAplic.FechaCamIva) Then
                If DBLet(RS.Fields(3).Value, "N") = vParamAplic.CodIvaGnral Then
                    CodIVA = vParamAplic.CodIvaGnralAnt
                ElseIf DBLet(RS.Fields(3).Value, "N") = vParamAplic.CodIvaRedu Then
                    CodIVA = vParamAplic.CodIvaReduAnt
                Else
                    CodIVA = DBLet(RS.Fields(3).Value, "N")
                End If
            Else
                CodIVA = DBLet(RS.Fields(3).Value, "N")
            End If
            
            Sql2 = ""
            Sql2 = DevuelveDesdeBDNew(cConta, "tiposiva", "porceiva", "codigiva", CodIVA, "N")
            If Sql2 = "" Then
                vPorcIva = 0
            Else
                vPorcIva = CCur(Sql2)
            End If
            
            ImporteSinIva = Round2(RS.Fields(6).Value / (1 + (vPorcIva / 100)), 2)
            If vParamAplic.Cooperativa = 2 Then ' si la cooperativa es el regaixo le quitamos el impuesto
                vImpuesto = Round2(DBLet(RS.Fields(5).Value, "N") * DBLet(RS.Fields(4).Value, "N"), 2)
                ImporteSinIva = ImporteSinIva - vImpuesto
            End If
            
            SQL1 = SQL1 & DBSet(ImporteSinIva, "N") & ")" ' importe sin iva
            
            Conn.Execute SQL1
        End If
        
        RS.MoveNext
    Wend
    
    RS.Close

eCargarTablaTemporal:
    Me.Pb1.visible = False
    Me.Label4(3).visible = False

    If Err.Number <> 0 Then
        MuestraError Err.Number, "Error en la carga de la tabla temporal"
    End If
End Sub

