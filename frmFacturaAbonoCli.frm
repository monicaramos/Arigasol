VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmFacturaAbonoCli 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Factura Abono Consumo a Clientes"
   ClientHeight    =   5955
   ClientLeft      =   45
   ClientTop       =   2430
   ClientWidth     =   6645
   Icon            =   "frmFacturaAbonoCli.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5955
   ScaleWidth      =   6645
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
      Height          =   5535
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   6375
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   6
         Left            =   1680
         MaxLength       =   10
         TabIndex        =   0
         Tag             =   "Código Postal|T|S|||clientes|codposta|||"
         Top             =   630
         Width           =   1050
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   5
         Left            =   2595
         Locked          =   -1  'True
         TabIndex        =   15
         Text            =   "Text5"
         Top             =   3285
         Width           =   3135
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   4
         Left            =   2595
         Locked          =   -1  'True
         TabIndex        =   14
         Text            =   "Text5"
         Top             =   2910
         Width           =   3165
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   5
         Left            =   1695
         MaxLength       =   3
         TabIndex        =   5
         Top             =   3285
         Width           =   830
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   4
         Left            =   1680
         MaxLength       =   3
         TabIndex        =   4
         Top             =   2910
         Width           =   830
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   3
         Left            =   1695
         MaxLength       =   10
         TabIndex        =   3
         Tag             =   "Código Postal|T|S|||clientes|codposta|||"
         Top             =   2430
         Width           =   1050
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   2
         Left            =   1695
         MaxLength       =   10
         TabIndex        =   2
         Tag             =   "Código Postal|T|S|||clientes|codposta|||"
         Top             =   2070
         Width           =   1050
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   4755
         TabIndex        =   7
         Top             =   4980
         Width           =   975
      End
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   3570
         TabIndex        =   6
         Top             =   4980
         Width           =   975
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   0
         Left            =   1665
         MaxLength       =   6
         TabIndex        =   1
         Top             =   1140
         Width           =   830
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   2565
         Locked          =   -1  'True
         TabIndex        =   9
         Text            =   "Text5"
         Top             =   1140
         Width           =   3135
      End
      Begin MSComctlLib.ProgressBar Pb1 
         Height          =   255
         Left            =   450
         TabIndex        =   20
         Top             =   4560
         Width           =   5250
         _ExtentX        =   9260
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   6
         Left            =   1380
         Picture         =   "frmFacturaAbonoCli.frx":000C
         ToolTipText     =   "Buscar fecha"
         Top             =   630
         Width           =   240
      End
      Begin VB.Label Label4 
         Caption         =   "Fecha Abono"
         ForeColor       =   &H00972E0B&
         Height          =   255
         Index           =   3
         Left            =   420
         TabIndex        =   19
         Top             =   390
         Width           =   1815
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   5
         Left            =   1395
         MouseIcon       =   "frmFacturaAbonoCli.frx":0097
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar colectivo"
         Top             =   3285
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   4
         Left            =   1410
         MouseIcon       =   "frmFacturaAbonoCli.frx":01E9
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar colectivo"
         Top             =   2910
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Grupo"
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
         Left            =   420
         TabIndex        =   18
         Top             =   2670
         Width           =   435
      End
      Begin VB.Label Label4 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   1
         Left            =   810
         TabIndex        =   17
         Top             =   3285
         Width           =   420
      End
      Begin VB.Label Label4 
         Caption         =   "Desde"
         Height          =   195
         Index           =   0
         Left            =   810
         TabIndex        =   16
         Top             =   2910
         Width           =   465
      End
      Begin VB.Label Label4 
         Caption         =   "Fecha Factura"
         ForeColor       =   &H00972E0B&
         Height          =   255
         Index           =   16
         Left            =   420
         TabIndex        =   13
         Top             =   1830
         Width           =   1815
      End
      Begin VB.Label Label4 
         Caption         =   "Desde"
         Height          =   195
         Index           =   15
         Left            =   810
         TabIndex        =   12
         Top             =   2070
         Width           =   465
      End
      Begin VB.Label Label4 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   14
         Left            =   810
         TabIndex        =   11
         Top             =   2430
         Width           =   420
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   2
         Left            =   1380
         Picture         =   "frmFacturaAbonoCli.frx":033B
         ToolTipText     =   "Buscar fecha"
         Top             =   2070
         Width           =   240
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   3
         Left            =   1380
         Picture         =   "frmFacturaAbonoCli.frx":03C6
         ToolTipText     =   "Buscar fecha"
         Top             =   2430
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
         Index           =   11
         Left            =   420
         TabIndex        =   10
         Top             =   900
         Width           =   660
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   0
         Left            =   1380
         MouseIcon       =   "frmFacturaAbonoCli.frx":0451
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar cliente"
         Top             =   1140
         Width           =   240
      End
   End
End
Attribute VB_Name = "frmFacturaAbonoCli"
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

Private WithEvents frmCol As frmManCoope 'Colectivo
Attribute frmCol.VB_VarHelpID = -1
Private WithEvents frmGru As frmManGrupo 'Grupos
Attribute frmGru.VB_VarHelpID = -1
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

Dim db As BaseDatos
Dim SerBonif As String ' letra de serie de las facturas de bonificacion
Dim SerGasol As String ' letra de serie de las facturas de gasolina

Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, 0, cerrar
    If cerrar Then Unload Me
End Sub


Private Sub cmdAceptar_Click()
Dim cDesde As String, cHasta As String 'cadena codigo Desde/Hasta
Dim nDesde As String, nHasta As String 'cadena Descripcion Desde/Hasta
Dim cadTabla As String, cOrden As String
Dim cadMen As String
Dim I As Byte
Dim SQL As String
Dim Tipo As Byte
Dim nRegs As Integer
Dim NumError As Long
Dim tipofac As String


'    MsgBox "En proceso de implementacion", vbExclamation
'    Exit Sub


    If Not DatosOk Then Exit Sub
    
    InicializarVbles
    MensError = ""
    '========= PARAMETROS  =============================
    'Añadir el parametro de Empresa
    cadParam = cadParam & "|pEmpresa=""" & vEmpresa.nomEmpre & """|"
    numParam = numParam + 1
    
     '======== FORMULA  ====================================
    'Seleccionar registros de la empresa conectada
'    Codigo = "{" & tabla & ".codempre}=" & vEmpresa.codEmpre
'    If Not AnyadirAFormula(cadFormula, Codigo) Then Exit Sub
'    If Not AnyadirAFormula(cadSelect, Codigo) Then Exit Sub
    
    
    'D/H Grupo
    cDesde = Trim(txtCodigo(4).Text)
    cHasta = Trim(txtCodigo(5).Text)
    nDesde = txtNombre(4).Text
    nHasta = txtNombre(5).Text
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{sligru.codempre}"
        TipCod = "N"
        If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHgrupo= """) Then Exit Sub
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
    
    ' Colectivo
    cDesde = Trim(txtCodigo(0).Text)
    nDesde = txtNombre(0).Text
    If Not (cDesde = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{ssocio.codcoope}"
        TipCod = "N"
        If Not PonerDesdeHasta(cDesde, "", nDesde, "", "pDHColec= """) Then Exit Sub
    End If
    
    SerBonif = ""
    SerBonif = DevuelveDesdeBDNew(cPTours, "stipom", "letraser", "codtipom", "FAB", "T")
    
    SerGasol = ""
    SerGasol = DevuelveDesdeBDNew(cPTours, "stipom", "letraser", "codtipom", "FAG", "T")
    
    
    'Comprobar si hay registros a procesar
    SQL = "select count(distinct sligru.codempre, slhfac.codartic, ssocio.tipsocio) "
    SQL = SQL & "from sligru, schfac, slhfac, ssocio, sartic, sfamia "
    SQL = SQL & " where sfamia.tipfamia = 1 " ' unicamente carburantes
    If txtCodigo(4).Text <> "" Then SQL = SQL & " and sligru.codempre >= " & DBSet(txtCodigo(4).Text, "N")
    If txtCodigo(5).Text <> "" Then SQL = SQL & " and sligru.codempre <= " & DBSet(txtCodigo(5).Text, "N")
    If txtCodigo(2).Text <> "" Then SQL = SQL & " and slhfac.fecfactu >= " & DBSet(txtCodigo(2).Text, "F")
    If txtCodigo(3).Text <> "" Then SQL = SQL & " and slhfac.fecfactu <= " & DBSet(txtCodigo(3).Text, "F")
    SQL = SQL & " and schfac.letraser = " & DBSet(SerGasol, "T")
'    sql = sql & " and schfac.letraser <> " & DBSet(SerBonif, "T")
    SQL = SQL & " and ssocio.codcoope = " & DBSet(txtCodigo(0).Text, "N")
    SQL = SQL & " and sligru.codsocio = schfac.codsocio "
    SQL = SQL & " and sligru.codsocio = ssocio.codsocio "
    SQL = SQL & " and sfamia.codfamia = sartic.codfamia "
    SQL = SQL & " and slhfac.codartic = sartic.codartic "
    SQL = SQL & " and slhfac.letraser = schfac.letraser and slhfac.numfactu = schfac.numfactu and slhfac.fecfactu = schfac.fecfactu "
    
    nRegs = TotalRegistros(SQL)
    
    If nRegs <> 0 Then
        ' cargamos el progresbar
        Pb1.visible = True
        CargarProgres Pb1, nRegs
        
        ' facturacion por cliente
        NumError = 0
        NumError = FacturaAbonoCli()
    Else
        MsgBox "No hay registros a procesar.", vbExclamation
        Exit Sub
    End If

eError:
    If Err.Number <> 0 Or NumError <> 0 Then
        MuestraError Err.Number, "No se ha realizado el proceso de abono a clientes. Llame a soporte." & vbCrLf & vbCrLf & _
              MensError
        db.RollbackTrans
        Set db = Nothing
        Pb1.visible = False
    Else
        db.CommitTrans
        MsgBox "Proceso finalizado correctamente", vbExclamation
        Set db = Nothing
        Pb1.visible = False
        cmdCancel_Click
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    If PrimeraVez Then
        PrimeraVez = False
        ValoresPorDefecto
        PonerFoco txtCodigo(6)
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
     Me.imgBuscar(4).Picture = frmPpal.imgListImages16.ListImages(1).Picture
     Me.imgBuscar(5).Picture = frmPpal.imgListImages16.ListImages(1).Picture

    '###Descomentar
'    CommitConexion
         
    FrameCobrosVisible True, h, w
    indFrame = 5
    tabla = "scaalb"
    
    Pb1.visible = False
    
    'Esto se consigue poneinedo el cancel en el opcion k corresponda
    Me.cmdCancel.Cancel = True
    Me.Width = w + 70
    Me.Height = h + 350
End Sub

Private Sub frmBpr_DatoSeleccionado(CadenaSeleccion As String)
'Form de Consulta de Colectivos
    txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "00")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    DesBloqueoManual ("FACTURAC")
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

Private Sub frmGru_DatoSeleccionado(CadenaSeleccion As String)
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

Private Sub ListView1_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub imgBuscar_Click(Index As Integer)
   Select Case Index
        Case 0 'COLECTIVO
            AbrirFrmColectivo (Index)
        
        Case 4, 5 'GRUPO
            AbrirFrmGrupo (Index)
        
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
    Select Case Index
        Case 0
            Me.Caption = "Facturación por Cliente"
        Case 1
            Me.Caption = "Facturación por Tarjeta"
        Case 2
            Me.Caption = "Facturación por Cliente y por Tarjeta"
    End Select
    
End Sub

Private Sub txtCodigo_GotFocus(Index As Integer)
    ConseguirFoco txtCodigo(Index), 3
End Sub

Private Sub txtCodigo_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    KEYdown KeyCode
End Sub

Private Sub txtCodigo_KeyPress(Index As Integer, KeyAscii As Integer)
' 15/02/2007
'   KEYpress KeyAscii
' ahora
    If KeyAscii = teclaBuscar Then
        Select Case Index
            Case 0: KEYBusqueda KeyAscii, 0 'colectivo
            Case 4: KEYBusqueda KeyAscii, 4 'grupo desde
            Case 5: KEYBusqueda KeyAscii, 5 'grupo hasta
            Case 6: KEYFecha KeyAscii, 6 'fecha abono
            Case 2: KEYFecha KeyAscii, 2 'fecha desde
            Case 3: KEYFecha KeyAscii, 3 'fecha hasta
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
            
        Case 0 'COLECTIVO
            txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), "scoope", "nomcoope", "codcoope", "N")
            If txtCodigo(Index).Text <> "" Then txtCodigo(Index).Text = Format(txtCodigo(Index).Text, "000")
        
        Case 2, 3, 6 'FECHAS
            If txtCodigo(Index).Text <> "" Then PonerFormatoFecha txtCodigo(Index)
        
        
        Case 4, 5 'GRUPO
            txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), "scagru", "nomempre", "codempre", "N")
            If txtCodigo(Index).Text <> "" Then txtCodigo(Index).Text = Format(txtCodigo(Index).Text, "000")
    
    End Select
End Sub

Private Sub FrameCobrosVisible(visible As Boolean, ByRef h As Integer, ByRef w As Integer)
    Me.FrameCobros.visible = visible
    If visible = True Then
        Me.FrameCobros.Top = -90
        Me.FrameCobros.Left = 0
        Me.FrameCobros.Height = 6015
        Me.FrameCobros.Width = 6555
        w = Me.FrameCobros.Width
        h = Me.FrameCobros.Height
    End If
End Sub

Private Sub ValoresPorDefecto()
    txtCodigo(6).Text = Format(Now, "dd/mm/yyyy")
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

Private Sub AbrirFrmGrupo(indice As Integer)
    indCodigo = indice
    Set frmGru = New frmManGrupo
    frmGru.DatosADevolverBusqueda = "0|1|"
    frmGru.DeConsulta = True
    frmGru.CodigoActual = txtCodigo(indCodigo)
    frmGru.Show vbModal
    Set frmGru = Nothing
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

Private Function PendientePasarTPV(SQL As String, Tipo As Byte) As Boolean
'Dim sql As String
Dim cadMen As String

    PendientePasarTPV = False
'    sql = "select count(*) from scaalb, ssocio, scoope where " & Formula & " and numfactu <> 0 and " & _
'          " scaalb.codsocio = ssocio.codsocio and ssocio.codcoope = scoope.codcoope "
    
    If Tipo <> 2 Then
        SQL = SQL & " and scoope.tipfactu = " & DBSet(Tipo, "N")
    Else 'VRS:2.0.2(1) añadida nueva opción
        SQL = SQL & " and (scoope.tipfactu = 0 or scoope.tipfactu = 1)"
    End If
    
    If (RegistrosAListar(SQL) <> 0) Then
        cadMen = "Hay registros pendientes de Traspaso a TPV." & vbCrLf & vbCrLf & _
                 "Debe realizar este proceso previamente." & vbCrLf & vbCrLf
        MsgBox cadMen, vbExclamation
        PendientePasarTPV = True
    End If
End Function

Private Function PendienteCierresTurno(DesFec As String, HasFec As String) As Boolean
Dim SQL As String
Dim cadMen As String

    PendienteCierresTurno = False
    SQL = "select count(*) from srecau where intconta = 0 "
    If DesFec <> "" Then SQL = SQL & " and fechatur >= " & DBSet(CDate(DesFec), "F") & " "
    If HasFec <> "" Then SQL = SQL & " and fechatur <= " & DBSet(CDate(HasFec), "F") & " "

    If (RegistrosAListar(SQL) <> 0) Then
        cadMen = "Quedan cierres de Turno por contabilizar. Revise." & vbCrLf & vbCrLf
        MsgBox cadMen, vbExclamation
        PendienteCierresTurno = True
    End If
    
End Function

Private Function DatosOk() As Boolean
Dim b As Boolean
Dim SQL As String
Dim Mens As String
Dim numfactu As String
Dim numser As String
Dim tipofac As String

    b = True

    If txtCodigo(6).Text = "" Then
        MsgBox "Debe introducir obligatoriamente una Fecha de Abono.", vbExclamation
        b = False
        PonerFoco txtCodigo(6)
    Else
        If Not FechaDentroPeriodoContable(CDate(txtCodigo(6).Text)) Then
            Mens = "La Fecha de Facturación no es del ejercicio actual ni siguiente. Reintroduzca."
            MsgBox Mens, vbExclamation
            b = False
            PonerFoco txtCodigo(6)
        Else
            '[Monica]20/06/2017: control de fechas que antes no estaba
            ResultadoFechaContaOK = EsFechaOKConta(CDate(txtCodigo(6)))
            If ResultadoFechaContaOK > 0 Then
                If ResultadoFechaContaOK <> 4 Then MsgBox MensajeFechaOkConta, vbExclamation
                b = False
                PonerFoco txtCodigo(6)
            Else
                'VRS:2.0.1(0)
                If Not FechaSuperiorUltimaLiquidacion(CDate(txtCodigo(6).Text)) Then
                    Mens = "  La Fecha de Facturación es inferior a la última liquidación de Iva. " & vbCrLf & vbCrLf
                    ' unicamente si el usuario es root el proceso continuará
                    If vSesion.Nivel > 0 Then
                        Mens = Mens & "  El proceso no continuará."
                        MsgBox Mens, vbExclamation
                        b = False
                        PonerFoco txtCodigo(6)
                    Else
                        Mens = Mens & "                        ¿ Desea continuar ?    " & vbCrLf
                        If MsgBox(Mens, vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
                            b = False
                            PonerFoco txtCodigo(6)
                        End If
                    End If
                End If
            End If
            
            ' la fecha de factura no debe ser inferior a la ultima factura de la serie
            numser = "letraser"
            numfactu = ""
            numfactu = DevuelveDesdeBDNew(cPTours, "stipom", "contador", "codtipom", "FAB", "T", numser)
            If numfactu <> "" Then
                If FechaFacturaInferiorUltimaFacturaSerieHco(CDate(txtCodigo(6).Text), CLng(numfactu), numser, 0) Then
                    Mens = "La Fecha de Factura es inferior a la última factura de la serie." & vbCrLf & vbCrLf
                    Mens = Mens & "                        ¿ Desea continuar ?    " & vbCrLf
                    If MsgBox(Mens, vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
                        b = False
                        PonerFoco txtCodigo(6)
                    End If
                End If
            End If
        End If
    End If
    
    ' comprobamos que el colectivo tenga de facturacion unicamente de cliente o tarjeta
    If b And txtCodigo(0).Text = "" Then
        Mens = "Debe introducir obligatoriamente un colectivo."
        MsgBox Mens, vbExclamation
        b = False
        PonerFoco txtCodigo(0)
    Else
        tipofac = ""
        tipofac = DevuelveDesdeBDNew(cPTours, "scoope", "tipfactu", "codcoope", txtCodigo(0).Text, "N")
        If tipofac = "2" Then ' el colectivo no tiene una facturacion por cliente o por tarjeta
            Mens = "Este colectivo no tiene tipo de facturación por cliente o tarjeta. Reintroduzca."
            MsgBox Mens, vbExclamation
            b = False
            PonerFoco txtCodigo(0)
        End If
    End If
    
    'comprobamos que haya articulo descuento en la tabla de parámetros.
    If b And vParamAplic.ArticDto = 0 Then
        MsgBox "Debe introducir un articulo de descuento en la tabla de parámetros. Revise.", vbExclamation
        b = False
        PonerFocoBtn cmdCancel
    Else
        'comprobamos que el articulo de descuento existe
        SQL = ""
        SQL = DevuelveDesdeBD("nomartic", "sartic", "codartic", vParamAplic.ArticDto, "N")
        If SQL = "" Then
            MsgBox "El artículo descuento de la tabla de parámetros no existe. Revise.", vbExclamation
            b = False
            PonerFocoBtn cmdCancel
        End If
    End If
    
    
    '03/03/2007 obligatoriamente ha de introducir fechadesde y fechahasta
    If b And (txtCodigo(2).Text = "" Or txtCodigo(3).Text = "") Then
        MsgBox "Debe introducir un rango de fechas obligatoriamente", vbExclamation
        b = False
        PonerFoco txtCodigo(2)
    End If
    
    DatosOk = b

End Function

Private Function FacturaAbonoCli() As Boolean
Dim SQL As String
Dim sql2 As String
Dim Rs As ADODB.Recordset
Dim ActCodempre As String
Dim ActCodartic As String
Dim ActTipsocio As String
Dim AntCodempre As String
Dim AntCodartic As String
Dim AntTipsocio As String
Dim HayReg As Byte
Dim v_linea As Integer
Dim NumError As Long
Dim BONIFICA As Currency

Dim nRegs As Integer

     On Error GoTo eFacturaAbonoCli


     Set db = New BaseDatos
     db.abrir vSesion.CadenaConexion, "root", "aritel"
     db.Tipo = "MYSQL"
     db.AbrirTrans

    NumError = 0

    ' realizamos la facturacion
    SQL = "select sligru.codempre, slhfac.codartic, ssocio.tipsocio, sum(cantidad) "
    SQL = SQL & "from sligru, schfac, slhfac, ssocio, sartic, sfamia "
    SQL = SQL & " where sfamia.tipfamia = 1 " ' unicamente carburantes
    If txtCodigo(4).Text <> "" Then SQL = SQL & " and sligru.codempre >= " & DBSet(txtCodigo(4).Text, "N")
    If txtCodigo(5).Text <> "" Then SQL = SQL & " and sligru.codempre <= " & DBSet(txtCodigo(5).Text, "N")
    If txtCodigo(2).Text <> "" Then SQL = SQL & " and slhfac.fecfactu >= " & DBSet(txtCodigo(2).Text, "F")
    If txtCodigo(3).Text <> "" Then SQL = SQL & " and slhfac.fecfactu <= " & DBSet(txtCodigo(3).Text, "F")
    SQL = SQL & " and schfac.letraser = " & DBSet(SerGasol, "T")
'    sql = sql & " and schfac.letraser <> " & DBSet(SerBonif, "T")
    SQL = SQL & " and ssocio.codcoope = " & DBSet(txtCodigo(0).Text, "N")
    SQL = SQL & " and sligru.codsocio = schfac.codsocio "
    SQL = SQL & " and sligru.codsocio = ssocio.codsocio "
    SQL = SQL & " and sfamia.codfamia = sartic.codfamia "
    SQL = SQL & " and slhfac.codartic = sartic.codartic "
    SQL = SQL & " and slhfac.letraser = schfac.letraser and slhfac.numfactu = schfac.numfactu and slhfac.fecfactu = schfac.fecfactu "
    SQL = SQL & " GROUP BY sligru.codempre, slhfac.codartic, ssocio.tipsocio "
    SQL = SQL & " ORDER BY sligru.codempre, slhfac.codartic, ssocio.tipsocio "

    Set Rs = db.cursor(SQL)
    HayReg = False
    v_linea = 0
    If Not Rs.EOF Then
        Rs.MoveFirst
        AntCodempre = Rs!codEmpre
'        AntCodartic = RS!codartic
'        AntTipsocio = RS!tipsocio

        HayReg = False


        NumError = BorramosTemporal(db)
        
        While Not Rs.EOF And NumError = 0
            HayReg = True
            IncrementarProgres Pb1, 1
            ActCodempre = Rs!codEmpre
'            ActCodartic = RS!codartic
'            ActTipsocio = RS!tipsocio
            If (ActCodempre <> AntCodempre) Then
                ' factura de abono para cliente
                NumError = FacturacionAbonoCliente(db, CCur(AntCodempre), txtCodigo(0).Text, txtCodigo(2).Text, txtCodigo(3).Text, txtCodigo(6).Text, SerGasol) '03/05/07 antes era serbonif
                'borramos temporal
                NumError = BorramosTemporal(db)
                AntCodempre = ActCodempre '03/05/07 añadida esta linea
            End If

            BONIFICA = Bonificacion(db, Rs.Fields(1).Value, Rs.Fields(3).Value, Rs.Fields(2).Value)
            If BONIFICA <> 0 Then
                SQL = "insert into tmpinformes (codusu, codigo1, precio1) values ( "
                SQL = SQL & vSesion.Codigo & "," & DBSet(Rs.Fields(1).Value, "N") & "," & DBSet(BONIFICA, "N") & ")"

                db.ejecutar SQL
            End If
            Rs.MoveNext

        Wend
        If HayReg And NumError = 0 Then
            'factura abono para cliente
            NumError = FacturacionAbonoCliente(db, CCur(ActCodempre), txtCodigo(0).Text, txtCodigo(2).Text, txtCodigo(3).Text, txtCodigo(6).Text, SerGasol) ' 03/05/07 antes era serbonif
            'borramos temporal
            If NumError = 0 Then NumError = BorramosTemporal(db)
        End If
    End If

eFacturaAbonoCli:
    If Err.Number <> 0 Or NumError <> 0 Then
        If Err.Number <> 0 Then
            FacturaAbonoCli = Err.Number
        Else
            FacturaAbonoCli = NumError
        End If
    Else
        FacturaAbonoCli = 0
    End If
End Function




Private Function Bonificacion(db As BaseDatos, articulo As String, cantidad As Currency, tipsocio As Byte) As Currency
Dim SQL As String
Dim Rs As ADODB.Recordset
Dim Encontrado As Boolean
Dim bonif As Currency

    SQL = "select numlinea, desdecan, hastacan, bonifica from sbonif "
    SQL = SQL & " where codartic = " & DBSet(articulo, "N") & " and tipsocio = " & DBSet(tipsocio, "N")
    SQL = SQL & " order by numlinea "
    
    Encontrado = False
    bonif = 0
    Set Rs = db.cursor(SQL)
    While Not Rs.EOF And Not Encontrado
        If Rs.Fields(1).Value <= cantidad And cantidad < Rs.Fields(2).Value Then
            Encontrado = True
            bonif = Rs.Fields(3).Value
        End If
        Rs.MoveNext
    Wend
        
    Bonificacion = bonif

    Set Rs = Nothing
    
End Function



