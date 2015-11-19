VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmEstRangos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Resumen Ventas por Rangos Horarios"
   ClientHeight    =   5715
   ClientLeft      =   45
   ClientTop       =   2430
   ClientWidth     =   6495
   Icon            =   "frmEstRangos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5715
   ScaleWidth      =   6495
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameCobros 
      Height          =   5625
      Left            =   30
      TabIndex        =   8
      Top             =   30
      Width           =   6435
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   4
         Left            =   1695
         MaxLength       =   2
         TabIndex        =   1
         Top             =   1200
         Width           =   830
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   5
         Left            =   1695
         MaxLength       =   2
         TabIndex        =   2
         Top             =   1575
         Width           =   830
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   4
         Left            =   2610
         Locked          =   -1  'True
         TabIndex        =   13
         Text            =   "Text5"
         Top             =   1200
         Width           =   3375
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   5
         Left            =   2610
         Locked          =   -1  'True
         TabIndex        =   12
         Text            =   "Text5"
         Top             =   1590
         Width           =   3375
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   6
         Left            =   1695
         MaxLength       =   3
         TabIndex        =   0
         Top             =   420
         Width           =   830
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   6
         Left            =   2625
         Locked          =   -1  'True
         TabIndex        =   11
         Text            =   "Text5"
         Top             =   420
         Width           =   3405
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   3
         Left            =   1695
         MaxLength       =   10
         TabIndex        =   4
         Tag             =   "Código Postal|T|S|||clientes|codposta|||"
         Top             =   2550
         Width           =   1050
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   2
         Left            =   1695
         MaxLength       =   10
         TabIndex        =   3
         Tag             =   "Código Postal|T|S|||clientes|codposta|||"
         Top             =   2190
         Width           =   1050
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   4980
         TabIndex        =   6
         Top             =   5010
         Width           =   975
      End
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   3930
         TabIndex        =   5
         Top             =   5010
         Width           =   975
      End
      Begin VB.Frame Frame4 
         Caption         =   "Tipo Informe"
         ForeColor       =   &H00972E0B&
         Height          =   1000
         Left            =   180
         TabIndex        =   7
         Top             =   3030
         Width           =   2175
         Begin VB.OptionButton Option1 
            Caption         =   "Solo Familias"
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   10
            Top             =   600
            Width           =   1815
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Detalle Artículos"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   9
            Top             =   240
            Value           =   -1  'True
            Width           =   1815
         End
      End
      Begin MSComctlLib.ProgressBar PB1 
         Height          =   225
         Left            =   210
         TabIndex        =   14
         Top             =   4500
         Width           =   5715
         _ExtentX        =   10081
         _ExtentY        =   397
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Label Label4 
         Caption         =   "Desde"
         Height          =   195
         Index           =   0
         Left            =   810
         TabIndex        =   22
         Top             =   1200
         Width           =   465
      End
      Begin VB.Label Label4 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   1
         Left            =   810
         TabIndex        =   21
         Top             =   1575
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
         Index           =   2
         Left            =   300
         TabIndex        =   20
         Top             =   960
         Width           =   480
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   4
         Left            =   1380
         MouseIcon       =   "frmEstRangos.frx":000C
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar familia"
         Top             =   1200
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   5
         Left            =   1380
         MouseIcon       =   "frmEstRangos.frx":015E
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar familia"
         Top             =   1575
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Rango Horario"
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
         Left            =   300
         TabIndex        =   19
         Top             =   420
         Width           =   1035
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   6
         Left            =   1380
         MouseIcon       =   "frmEstRangos.frx":02B0
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar colectivo"
         Top             =   420
         Width           =   240
      End
      Begin VB.Label Label4 
         Caption         =   "Fecha "
         ForeColor       =   &H00972E0B&
         Height          =   255
         Index           =   16
         Left            =   300
         TabIndex        =   18
         Top             =   1950
         Width           =   1815
      End
      Begin VB.Label Label4 
         Caption         =   "Desde"
         Height          =   195
         Index           =   15
         Left            =   810
         TabIndex        =   17
         Top             =   2190
         Width           =   465
      End
      Begin VB.Label Label4 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   14
         Left            =   810
         TabIndex        =   16
         Top             =   2550
         Width           =   420
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   2
         Left            =   1380
         Picture         =   "frmEstRangos.frx":0402
         ToolTipText     =   "Buscar fecha"
         Top             =   2190
         Width           =   240
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   3
         Left            =   1380
         Picture         =   "frmEstRangos.frx":048D
         ToolTipText     =   "Buscar fecha"
         Top             =   2550
         Width           =   240
      End
      Begin VB.Label Label4 
         Caption         =   "Proceso:"
         Height          =   195
         Index           =   4
         Left            =   270
         TabIndex        =   15
         Top             =   4230
         Width           =   3405
      End
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   7680
      Top             =   5160
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmEstRangos"
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

Private WithEvents frmC As frmCal 'Calendario
Attribute frmC.VB_VarHelpID = -1
Private WithEvents frmRan As frmManRangos 'Rangos Horarios
Attribute frmRan.VB_VarHelpID = -1
Private WithEvents frmFam As frmManFamia 'Familias
Attribute frmFam.VB_VarHelpID = -1

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
Dim Sql As String
Dim TotalReg As Integer
Dim Rs As ADODB.Recordset
Dim NumCli As Long

InicializarVbles
    
    If txtCodigo(6).Text = "" Then
        MsgBox "Introduzca el Rango Horario a listar.", vbExclamation
        Exit Sub
    End If
    
    '========= PARAMETROS  =============================
    'Añadir el parametro de Empresa
    cadParam = cadParam & "|pEmpresa=""" & vEmpresa.nomEmpre & """|"
    numParam = numParam + 1
    
    ' miramos en el historico slhfac
    Sql = "select distinct ((time(linrangos.desdehora)<=  time(slhfac.horalbar)) and ( time(slhfac.horalbar)<= time(linrangos.hastahora))) from slhfac, linrangos  where 1=1 "
    If txtCodigo(2).Text <> "" Then Sql = Sql & " and fecalbar >= " & DBSet(txtCodigo(2), "F")
    If txtCodigo(3).Text <> "" Then Sql = Sql & " and fecalbar <= " & DBSet(txtCodigo(3), "F")
    Sql = Sql & " and linrangos.codigo = " & DBSet(txtCodigo(6).Text, "N")
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    TotalReg = 0
    If Not Rs.EOF Then
        If Rs.Fields(0).Value > 0 Then
            TotalReg = 1
        Else
            Rs.MoveNext
            If Not Rs.EOF Then TotalReg = 1  'Solo es para saber que hay registros que mostrar
        End If
    End If
    Rs.Close
    Set Rs = Nothing
    
    ' miramos en scaalb
    Sql = "select distinct ((time(linrangos.desdehora)<=  time(scaalb.horalbar)) and ( time(scaalb.horalbar)<= time(linrangos.hastahora))) from scaalb, linrangos  where 1=1 "
    If txtCodigo(2).Text <> "" Then Sql = Sql & " and fecalbar >= " & DBSet(txtCodigo(2), "F")
    If txtCodigo(3).Text <> "" Then Sql = Sql & " and fecalbar <= " & DBSet(txtCodigo(3), "F")
    Sql = Sql & " and linrangos.codigo = " & DBSet(txtCodigo(6).Text, "N")
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    If Not Rs.EOF Then
        If Rs.Fields(0).Value > 0 Then
            TotalReg = 1
        Else
            Rs.MoveNext
            If Not Rs.EOF Then TotalReg = 1  'Solo es para saber que hay registros que mostrar
        End If
    End If
    Rs.Close
    Set Rs = Nothing
    
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
    
    'D/H Fecha
    cDesde = Trim(txtCodigo(2).Text)
    cHasta = Trim(txtCodigo(3).Text)
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{" & tabla & ".fecalbar}"
        TipCod = "F"
        If Not PonerDesdeHasta(cDesde, cHasta, "", "", "pDHfechaFac= """) Then Exit Sub
    End If
    
    
    If TotalReg > 0 Then
          BorradoTablaIntermedia
          'CargarPb1
          Me.Label4(4).visible = True
          CargarTablaIntermedia ("slhfac")
          CargarTablaIntermedia ("scaalb")
          CargarTotalClientes
          Me.Pb1.visible = False
          Me.Label4(4).visible = False
          cadFormula = "{cabrangos.codigo} = {@pCodigo} and {tmpinformes.codusu} = " & vSesion.Codigo
          cadParam = cadParam & "|pCodigo= " & txtCodigo(6).Text
          numParam = numParam + 1
          If Option1(0).Value = True Then
            cadParam = cadParam & "|pDetalle= 1"
          Else
            cadParam = cadParam & "|pDetalle= 0"
          End If
          numParam = numParam + 1
          
          'pcodigo me indica si hemos puesto la misma familia desde y hasta
          If txtCodigo(4).Text <> txtCodigo(5).Text Then
              cadParam = cadParam & "|pFamUnica= 0|"
          Else
              cadParam = cadParam & "|pFamUnica= 1|"
          End If
          numParam = numParam + 1
          
          cadTitulo = "Estadísticas de Ventas Artículos por Rango Horario"
          cadNombreRPT = "rEstRango.rpt"
          LlamarImprimir
    Else
        MsgBox "No existen datos en este período. Reintroduzca", vbExclamation
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
    Limpiar Me

    'IMAGES para busqueda
     Me.imgBuscar(6).Picture = frmPpal.imgListImages16.ListImages(1).Picture
     Me.imgBuscar(4).Picture = frmPpal.imgListImages16.ListImages(1).Picture
     Me.imgBuscar(5).Picture = frmPpal.imgListImages16.ListImages(1).Picture

    '###Descomentar
'    CommitConexion
         
    FrameCobrosVisible True, h, w
    indFrame = 5
    tabla = "slhfac"
            
    Me.Pb1.visible = False
    Me.Label4(4).visible = False
    
    'Esto se consigue poneinedo el cancel en el opcion k corresponda
    Me.cmdCancel.Cancel = True
End Sub

Private Sub frmC_Selec(vFecha As Date)
 'Fecha
    txtCodigo(CByte(imgFec(2).Tag)).Text = Format(vFecha, "dd/MM/yyyy")
End Sub

Private Sub frmFam_DatoSeleccionado(CadenaSeleccion As String)
    txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "00")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmRan_DatoSeleccionado(CadenaSeleccion As String)
    txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000")
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
        Case 4, 5 'FAMILIAS
            AbrirFrmFamilias (Index)
        
        Case 6 'COLECTIVO
            AbrirFrmRangos (Index)
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
            Case 6: KEYBusqueda KeyAscii, 6 'rango horario
            Case 4: KEYBusqueda KeyAscii, 4 'familia desde
            Case 5: KEYBusqueda KeyAscii, 5 'familia hasta
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
Dim cad As String, cadTipo As String 'tipo cliente

    'Quitar espacios en blanco por los lados
    txtCodigo(Index).Text = Trim(txtCodigo(Index).Text)
    
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub

    
    Select Case Index
        Case 2, 3 'FECHAS
            If txtCodigo(Index).Text <> "" Then PonerFormatoFecha txtCodigo(Index)
        
        Case 4, 5 'FAMILIAS
            txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), "sfamia", "nomfamia", "codfamia", "N")
            If txtCodigo(Index).Text <> "" Then txtCodigo(Index).Text = Format(txtCodigo(Index).Text, "00")
        
        Case 6 'COLECTIVOS
            txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), "cabrangos", "descripcion", "codigo", "N")
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
        .Opcion = 0
        .Show vbModal
    End With
End Sub

Private Sub AbrirFrmRangos(indice As Integer)
    indCodigo = indice
    Set frmRan = New frmManRangos
    frmRan.DatosADevolverBusqueda = "0|1|"
    frmRan.DeConsulta = True
    frmRan.CodigoActual = txtCodigo(indCodigo)
    frmRan.Show vbModal
    Set frmRan = Nothing
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

Private Sub CargarTablaIntermedia(tabla As String)
Dim Rs As ADODB.Recordset
Dim Rs1 As ADODB.Recordset
Dim Sql As String
Dim Sql2 As String
Dim Sql3 As String
Dim Fam As String
Dim NumLin As Integer

    CargarPb1 (tabla)
    
    If tabla = "slhfac" Then ' slhfac
        Sql = "select horalbar, codartic, cantidad, implinea from " & tabla & " where 1 = 1"
    Else ' scaalb
        Sql = "select horalbar, codartic, cantidad, importel from " & tabla & " where 1 = 1"
    End If
    
    If txtCodigo(2).Text <> "" Then Sql = Sql & " and fecalbar >= " & DBSet(txtCodigo(2).Text, "F")
    If txtCodigo(3).Text <> "" Then Sql = Sql & " and fecalbar <= " & DBSet(txtCodigo(3).Text, "F")
    
    If txtCodigo(4).Text <> "" Or txtCodigo(5) <> "" Then
        Sql = Sql & " and codartic in (select codartic from sartic where 1=1 "
        If txtCodigo(4).Text <> "" Then Sql = Sql & " and codfamia >= " & DBSet(txtCodigo(4).Text, "N")
        If txtCodigo(5).Text <> "" Then Sql = Sql & " and codfamia <= " & DBSet(txtCodigo(5).Text, "N")
        Sql = Sql & ")"
    End If
    
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    While Not Rs.EOF
        Sql2 = "select numlinea from linrangos where time(desdehora) <= " & DBSet(Rs!horalbar, "H")
        Sql2 = Sql2 & " and time(hastahora) >= " & DBSet(Rs!horalbar, "H")
        Sql2 = Sql2 & " and codigo = " & DBSet(txtCodigo(6).Text, "N")
        
        Set Rs1 = New ADODB.Recordset
        Rs1.Open Sql2, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        NumLin = 99
        If Not Rs1.EOF Then
            NumLin = Rs1.Fields(0).Value
        End If
        
        Rs1.Close
        Set Rs1 = Nothing
        
        Sql2 = "select * from tmpinformes where codusu = " & vSesion.Codigo & " and codigo1 = " & DBSet(NumLin, "N")
        Sql2 = Sql2 & " and campo2 = " & DBSet(Rs!codArtic, "N")
        Set Rs1 = New ADODB.Recordset
        Rs1.Open Sql2, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
        If Not Rs1.EOF Then
            ' modificamos el registro correspondiente
            Sql3 = "update tmpinformes set importe2 = importe2 + " & DBSet(Rs.Fields(3).Value, "N") & ","
            Sql3 = Sql3 & "importe3 = importe3 + " & DBSet(Rs!cantidad, "N") & " where codusu = " & DBSet(vSesion.Codigo, "N")
            Sql3 = Sql3 & " and codigo1 = " & DBSet(NumLin, "N") & " and campo2 = " & DBSet(Rs!codArtic, "N")
            
            Conn.Execute Sql3
        Else
            ' insertamos el nuevo registro en la tabla intermedia
            Fam = ""
            Fam = DevuelveDesdeBD("codfamia", "sartic", "codartic", Rs!codArtic, "N")
            
            Sql3 = "insert into tmpinformes (codusu, codigo1, campo1, campo2, importe1, importe2, importe3, importe4, importe5) values ("
            Sql3 = Sql3 & vSesion.Codigo & "," & DBSet(NumLin, "N") & "," & DBSet(Fam, "N") & "," & DBSet(Rs!codArtic, "N") & ","
            Sql3 = Sql3 & "0," & DBSet(Rs.Fields(3).Value, "N") & "," & DBSet(Rs!cantidad, "N") & ",0,0)"
            
            Conn.Execute Sql3
        End If
        
        Pb1.Value = Pb1.Value + 1
        
        Rs1.Close
        Set Rs1 = Nothing
        
        Rs.MoveNext
    Wend
    
    Rs.Close
    Set Rs = Nothing
End Sub


Private Sub BorradoTablaIntermedia()
Dim Sql As String

    Sql = "delete from tmpinformes where codusu = " & vSesion.Codigo
    Conn.Execute Sql
    
End Sub

Private Sub CargarPb1(tabla As String)
Dim Sql As String
Dim NumLinea As Long

    If tabla = "slhfac" Then
        Me.Label4(4).Caption = "Procesando tabla de histórico de facturas. "
    Else
        Me.Label4(4).Caption = "Procesando tabla de albaranes. "
    End If

    Me.Label4(4).Refresh
    
    Sql = "select count(*) from " & tabla & " where 1 = 1 "
    
    If txtCodigo(2).Text <> "" Then Sql = Sql & " and fecalbar >= " & DBSet(txtCodigo(2).Text, "F")
    If txtCodigo(3).Text <> "" Then Sql = Sql & " and fecalbar <= " & DBSet(txtCodigo(3).Text, "F")
    
    If txtCodigo(4).Text <> "" Or txtCodigo(5) <> "" Then
        Sql = Sql & " and codartic in (select codartic from sartic where 1=1 "
        If txtCodigo(4).Text <> "" Then Sql = Sql & " and codfamia >= " & DBSet(txtCodigo(4).Text, "N")
        If txtCodigo(5).Text <> "" Then Sql = Sql & " and codfamia <= " & DBSet(txtCodigo(5).Text, "N")
        Sql = Sql & ")"
    End If
    
    NumLinea = TotalRegistros(Sql) + 1
    
'    SQL = "select count(*) from scaalb where 1 = 1 "
'
'    If txtCodigo(2).Text <> "" Then SQL = SQL & " and fecalbar >= " & DBSet(txtCodigo(2).Text, "F")
'    If txtCodigo(3).Text <> "" Then SQL = SQL & " and fecalbar <= " & DBSet(txtCodigo(3).Text, "F")
'
'    If txtCodigo(4).Text <> "" Or txtCodigo(5) <> "" Then
'        SQL = SQL & " and codartic in (select codartic from sartic where 1=1 "
'        If txtCodigo(4).Text <> "" Then SQL = SQL & " and codfamia >= " & DBSet(txtCodigo(4).Text, "N")
'        If txtCodigo(5).Text <> "" Then SQL = SQL & " and codfamia <= " & DBSet(txtCodigo(5).Text, "N")
'        SQL = SQL & ")"
'    End If
'
'    NumLinea = NumLinea + TotalRegistros(SQL) + 1

    If NumLinea > 0 Then
        Me.Pb1.visible = True
        Pb1.Max = NumLinea
        Pb1.Value = 0
        
    End If
End Sub

Private Sub CargarTotalClientes()
Dim Sql As String
Dim Sql2 As String
Dim Rs As ADODB.Recordset
Dim NumLinea As Long
Dim NumCli As Long

    Label4(4).Caption = "Procesando el Total por Clientes. "
    Me.Label4(4).Refresh

    Sql = "select count(*) from tmpinformes where codusu = " & DBSet(vSesion.Codigo, "N")
    
    NumLinea = TotalRegistros(Sql) + 1
    If NumLinea > 0 Then
        Me.Pb1.visible = True
        Pb1.Max = NumLinea
        Pb1.Value = 0
    End If
    
    Sql = "select campo2, codigo1, campo1 from tmpinformes where codusu = " & DBSet(vSesion.Codigo, "N")
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    
    Conn.Execute " DROP TABLE IF EXISTS tmpsocio;"
    Sql = "CREATE TEMPORARY TABLE tmpsocio ( "
    Sql = Sql & "codsocio mediumint(7) unsigned NOT NULL default '0')"
    Conn.Execute Sql
     
    Conn.Execute " DROP TABLE IF EXISTS tmpsocio2;"
    Sql = "CREATE TEMPORARY TABLE tmpsocio2 ( "
    Sql = Sql & "codsocio mediumint(7) unsigned NOT NULL default '0')"
    Conn.Execute Sql
     
     
     
    While Not Rs.EOF
            ' TOTAL POR ARTICULO
            If Option1(0).Value = True Then
                    Sql2 = "insert into tmpsocio select distinct codsocio from schfac, slhfac "
                    Sql2 = Sql2 & " where slhfac.codartic = " & DBSet(Rs!campo2, "N")
                           
                    If txtCodigo(2).Text <> "" Then Sql2 = Sql2 & " and slhfac.fecalbar >= " & DBSet(txtCodigo(2).Text, "F")
                    If txtCodigo(3).Text <> "" Then Sql2 = Sql2 & " and slhfac.fecalbar <= " & DBSet(txtCodigo(3).Text, "F")
                    
                    Sql2 = Sql2 & " and (select time(desdehora) <= time (slhfac.horalbar ) and time (slhfac.horalbar) <= time(hastahora) "
                    Sql2 = Sql2 & " from linrangos where numlinea = " & DBSet(Rs!Codigo1, "N") & " and codigo = " & DBSet(txtCodigo(6).Text, "N") & ")"
                    Sql2 = Sql2 & " and schfac.numfactu = slhfac.numfactu and schfac.letraser = slhfac.letraser and schfac.fecfactu = slhfac.fecfactu "
                    
                    ' añadido
                    Conn.Execute Sql2
        
                    Sql2 = "insert into tmpsocio select distinct codsocio from scaalb "
                    Sql2 = Sql2 & " where scaalb.codartic = " & DBSet(Rs!campo2, "N")
                           
                    If txtCodigo(2).Text <> "" Then Sql2 = Sql2 & " and scaalb.fecalbar >= " & DBSet(txtCodigo(2).Text, "F")
                    If txtCodigo(3).Text <> "" Then Sql2 = Sql2 & " and scaalb.fecalbar <= " & DBSet(txtCodigo(3).Text, "F")
                    
                    Sql2 = Sql2 & " and (select time(desdehora) <= time (scaalb.horalbar ) and time (scaalb.horalbar) <= time(hastahora) "
                    Sql2 = Sql2 & " from linrangos where numlinea = " & DBSet(Rs!Codigo1, "N") & " and codigo = " & DBSet(txtCodigo(6).Text, "N") & ")"
                    
                    Conn.Execute Sql2
                    
                    'añadidas
                    Sql2 = "select count(distinct codsocio) from tmpsocio "
                    NumCli = TotalRegistros(Sql2)
                    
                    
                    ' actualizamos el registro correspondiente de tmpinformes
                    Sql2 = "update tmpinformes set importe1 =  " & DBSet(NumCli, "N") & _
                           " where codusu = " & vSesion.Codigo & _
                           " and codigo1 = " & DBSet(Rs!Codigo1, "N") & _
                           " and campo2 = " & DBSet(Rs!campo2, "N")
                    Conn.Execute Sql2
            
                    Sql2 = "delete from tmpsocio"
                    Conn.Execute Sql2
            End If
    
            ' TOTAL POR FAMILIA
            Sql2 = "insert  into tmpsocio select distinct codsocio from schfac, slhfac, sartic "
            Sql2 = Sql2 & " where sartic.codfamia = " & DBSet(Rs!campo1, "N")
                   
            If txtCodigo(2).Text <> "" Then Sql2 = Sql2 & " and slhfac.fecalbar >= " & DBSet(txtCodigo(2).Text, "F")
            If txtCodigo(3).Text <> "" Then Sql2 = Sql2 & " and slhfac.fecalbar <= " & DBSet(txtCodigo(3).Text, "F")
            
            Sql2 = Sql2 & " and (select time(desdehora) <= time (slhfac.horalbar ) and time (slhfac.horalbar) <= time(hastahora) "
            Sql2 = Sql2 & " from linrangos where numlinea = " & DBSet(Rs!Codigo1, "N") & " and codigo = " & DBSet(txtCodigo(6).Text, "N") & ")"
            Sql2 = Sql2 & " and schfac.numfactu = slhfac.numfactu and schfac.letraser = slhfac.letraser and schfac.fecfactu = slhfac.fecfactu "
            Sql2 = Sql2 & " and slhfac.codartic = sartic.codartic "
            
            Conn.Execute Sql2
            Conn.Execute "insert into tmpsocio2 select * from tmpsocio"
            
            Sql2 = "insert into tmpsocio select distinct codsocio from scaalb, sartic "
            Sql2 = Sql2 & " where sartic.codfamia = " & DBSet(Rs!campo1, "N")
                   
            If txtCodigo(2).Text <> "" Then Sql2 = Sql2 & " and scaalb.fecalbar >= " & DBSet(txtCodigo(2).Text, "F")
            If txtCodigo(3).Text <> "" Then Sql2 = Sql2 & " and scaalb.fecalbar <= " & DBSet(txtCodigo(3).Text, "F")
            
            Sql2 = Sql2 & " and (select time(desdehora) <= time (scaalb.horalbar ) and time (scaalb.horalbar) <= time(hastahora) "
            Sql2 = Sql2 & " from linrangos where numlinea = " & DBSet(Rs!Codigo1, "N") & " and codigo = " & DBSet(txtCodigo(6).Text, "N") & ")"
            Sql2 = Sql2 & " and scaalb.codartic = sartic.codartic"
            
            Conn.Execute Sql2
            Conn.Execute "insert into tmpsocio2 select * from tmpsocio"
            
            'añadidas
            Sql2 = "select count(distinct codsocio) from tmpsocio "
            NumCli = TotalRegistros(Sql2)
                       
            
            ' actualizamos el registro correspondiente de tmpinformes
            Sql2 = "update tmpinformes set importe4 = " & DBSet(NumCli, "N") & _
                   " where codusu = " & vSesion.Codigo & _
                   " and codigo1 = " & DBSet(Rs!Codigo1, "N") & _
                   " and campo1 = " & DBSet(Rs!campo1, "N")
            Conn.Execute Sql2
            Sql2 = "delete from tmpsocio"
            Conn.Execute Sql2
        
            
            ' TOTAL GLOBAL
            Sql2 = "insert into tmpsocio select distinct codsocio from schfac, slhfac "
            Sql2 = Sql2 & " where 1 = 1 "

            If txtCodigo(2).Text <> "" Then Sql2 = Sql2 & " and slhfac.fecalbar >= " & DBSet(txtCodigo(2).Text, "F")
            If txtCodigo(3).Text <> "" Then Sql2 = Sql2 & " and slhfac.fecalbar <= " & DBSet(txtCodigo(3).Text, "F")

            Sql2 = Sql2 & " and (select time(desdehora) <= time (slhfac.horalbar ) and time (slhfac.horalbar) <= time(hastahora) "
            Sql2 = Sql2 & " from linrangos where numlinea = " & DBSet(Rs!Codigo1, "N") & " and codigo = " & DBSet(txtCodigo(6).Text, "N") & ")"
            Sql2 = Sql2 & " and schfac.numfactu = slhfac.numfactu and schfac.letraser = slhfac.letraser and schfac.fecfactu = slhfac.fecfactu "

            Conn.Execute Sql2

            Sql2 = "insert into tmpsocio select distinct codsocio from scaalb "
            Sql2 = Sql2 & " where 1=1 "

            If txtCodigo(2).Text <> "" Then Sql2 = Sql2 & " and scaalb.fecalbar >= " & DBSet(txtCodigo(2).Text, "F")
            If txtCodigo(3).Text <> "" Then Sql2 = Sql2 & " and scaalb.fecalbar <= " & DBSet(txtCodigo(3).Text, "F")

            Sql2 = Sql2 & " and (select time(desdehora) <= time (scaalb.horalbar ) and time (scaalb.horalbar) <= time(hastahora) "
            Sql2 = Sql2 & " from linrangos where numlinea = " & DBSet(Rs!Codigo1, "N") & " and codigo = " & DBSet(txtCodigo(6).Text, "N") & ")"

            Conn.Execute Sql2

            'añadidas
            Sql2 = "select count(distinct codsocio) from tmpsocio "
            NumCli = TotalRegistros(Sql2)

            ' actualizamos el registro correspondiente de tmpinformes
            Sql2 = "update tmpinformes set importe5 =  " & DBSet(NumCli, "N") & _
                   " where codusu = " & vSesion.Codigo & _
                   " and codigo1 = " & DBSet(Rs!Codigo1, "N") '& _
'                   " and campo2 = " & DBSet(RS!campo2, "N")
            Conn.Execute Sql2
            Sql2 = "delete from tmpsocio"
            Conn.Execute Sql2

        
        'incrementamos el valor del progressbar
        Pb1.Value = Pb1.Value + 1
        
        Rs.MoveNext
    Wend
    
    Label4(4).Caption = "Procesando el Total Global "
    Me.Label4(4).Refresh
    
    NumCli = TotalRegistros("select count(distinct codsocio) from tmpsocio2")
    Sql2 = "update tmpinformes set importeb1 = " & DBSet(NumCli, "N") & " where codusu =" & vSesion.Codigo
    Conn.Execute Sql2
    
    Conn.Execute " DROP TABLE IF EXISTS tmpsocio;"
    Conn.Execute " DROP TABLE IF EXISTS tmpsocio2;"
    
    Rs.Close
    Set Rs = Nothing
End Sub

