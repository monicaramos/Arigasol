VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmEvoMensCli 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Evolución Mensual de Clientes"
   ClientHeight    =   4560
   ClientLeft      =   45
   ClientTop       =   2430
   ClientWidth     =   7185
   Icon            =   "frmEvoMensCli.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4560
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
      Height          =   4335
      Left            =   120
      TabIndex        =   4
      Top             =   150
      Width           =   6915
      Begin VB.Frame Frame4 
         Caption         =   "Orden"
         ForeColor       =   &H00972E0B&
         Height          =   765
         Left            =   480
         TabIndex        =   14
         Top             =   2190
         Width           =   5925
         Begin VB.OptionButton Option1 
            Caption         =   "Mes Pedido"
            Height          =   255
            Index           =   2
            Left            =   3900
            TabIndex        =   17
            Top             =   270
            Width           =   1515
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Alfabético"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   16
            Top             =   270
            Value           =   -1  'True
            Width           =   1515
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Acumulado"
            Height          =   255
            Index           =   1
            Left            =   2070
            TabIndex        =   15
            Top             =   270
            Width           =   1515
         End
      End
      Begin VB.ComboBox Combo1 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   1
         ItemData        =   "frmEvoMensCli.frx":000C
         Left            =   3780
         List            =   "frmEvoMensCli.frx":000E
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Tag             =   "Tipo Familia|N|N|0|9|sfamia|tipfamia|||"
         Top             =   1680
         Width           =   1560
      End
      Begin VB.ComboBox Combo1 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   0
         ItemData        =   "frmEvoMensCli.frx":0010
         Left            =   1320
         List            =   "frmEvoMensCli.frx":0012
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Tag             =   "Tipo Familia|N|N|0|9|sfamia|tipfamia|||"
         Top             =   1680
         Width           =   1560
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   4905
         TabIndex        =   3
         Top             =   3600
         Width           =   975
      End
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   3720
         TabIndex        =   2
         Top             =   3600
         Width           =   975
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   0
         Left            =   1830
         MaxLength       =   6
         TabIndex        =   0
         Top             =   840
         Width           =   1065
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   1
         Left            =   1830
         MaxLength       =   6
         TabIndex        =   1
         Top             =   1215
         Width           =   1035
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   3090
         Locked          =   -1  'True
         TabIndex        =   6
         Text            =   "Text5"
         Top             =   840
         Width           =   3135
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   3090
         Locked          =   -1  'True
         TabIndex        =   5
         Text            =   "Text5"
         Top             =   1200
         Width           =   3135
      End
      Begin MSComctlLib.ProgressBar Pb1 
         Height          =   255
         Left            =   480
         TabIndex        =   18
         Top             =   3180
         Width           =   5880
         _ExtentX        =   10372
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Label Label4 
         Caption         =   "Año"
         ForeColor       =   &H00972E0B&
         Height          =   255
         Index           =   0
         Left            =   3120
         TabIndex        =   11
         Top             =   1740
         Width           =   1815
      End
      Begin VB.Label Label4 
         Caption         =   "Mes"
         ForeColor       =   &H00972E0B&
         Height          =   255
         Index           =   16
         Left            =   630
         TabIndex        =   10
         Top             =   1740
         Width           =   1815
      End
      Begin VB.Label Label4 
         Caption         =   "Desde"
         Height          =   195
         Index           =   13
         Left            =   960
         TabIndex        =   9
         Top             =   840
         Width           =   465
      End
      Begin VB.Label Label4 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   12
         Left            =   960
         TabIndex        =   8
         Top             =   1215
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
         TabIndex        =   7
         Top             =   600
         Width           =   495
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   0
         Left            =   1560
         MouseIcon       =   "frmEvoMensCli.frx":0014
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar familia"
         Top             =   840
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   1
         Left            =   1530
         MouseIcon       =   "frmEvoMensCli.frx":0166
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar familia"
         Top             =   1215
         Width           =   240
      End
   End
End
Attribute VB_Name = "frmEvoMensCli"
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

Private WithEvents frmCli As frmManClien 'Clientes
Attribute frmCli.VB_VarHelpID = -1
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

Dim mes1 As Integer
Dim mes2 As Integer
Dim mes3 As Integer
Dim mes4 As Integer
Dim mes5 As Integer

Dim ano1 As Integer
Dim ano2 As Integer
Dim ano3 As Integer
Dim ano4 As Integer
Dim ano5 As Integer

Dim Opcion As Integer
Dim nRegs As Integer

Dim desfec As String
Dim hasfec As String


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
Dim sql As String

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
        Codigo = "{schfac.codsocio}"
        TipCod = "N"
        If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHClie= """) Then Exit Sub
    End If
    Opcion = 0
    If Option1(1).Value Then
        cadParam = cadParam & "pOrden= {tmpinformes.importe6}|"
        Opcion = 100
    End If
    If Option1(2).Value Then
        cadParam = cadParam & "pOrden= {tmpinformes.importe5}|"
        Opcion = 100
    End If
    numParam = numParam + 1
    
    cadTABLA = tabla
    
    
    CargarParametros
    
    
    sql = "select count(distinct codsocio) from schfac where fecfactu between " & DBSet(desfec, "F") & " and " & DBSet(hasfec, "F")
    
    If txtCodigo(0).Text <> "" Then sql = sql & " and codsocio >= " & DBSet(txtCodigo(0).Text, "N")
    If txtCodigo(1).Text <> "" Then sql = sql & " and codsocio <= " & DBSet(txtCodigo(1).Text, "N")
    nRegs = TotalRegistros(sql)
    
    
    If nRegs <> 0 Then
       CargarTemporal
       cadTitulo = "Evolución Mensual Clientes"
       cadNombreRPT = "rEvolCli.rpt"
       LlamarImprimir
       'AbrirVisReport
    Else
        MsgBox "No hay datos entre estos límites. Reintroduzca", vbExclamation
        PonerFoco txtCodigo(0)
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
Dim mes  As Integer
    PrimeraVez = True
    Limpiar Me

    'IMAGES para busqueda
     Me.imgBuscar(0).Picture = frmPpal.imgListImages16.ListImages(1).Picture
     Me.imgBuscar(1).Picture = frmPpal.imgListImages16.ListImages(1).Picture

    '###Descomentar
'    CommitConexion
         
    FrameCobrosVisible True, h, w
    indFrame = 5
    tabla = "schfac"
            
    CargaCombo
    
    mes = Month(Now) - 1
    If mes = 0 Then mes = 12
    
    PosicionarCombo Combo1(0), mes
    Combo1(1).Text = Format(Year(Now), "####")
    
    'Esto se consigue poneinedo el cancel en el opcion k corresponda
    Me.cmdCancel.Cancel = True
    Pb1.visible = False
'    Me.Width = w + 70
'    Me.Height = h + 350
End Sub


Private Sub frmFam_DatoSeleccionado(CadenaSeleccion As String)
'Form de Consulta de Familias
    txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "00")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub ListView1_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub imgBuscar_Click(Index As Integer)
   Select Case Index
        
        Case 0, 1 'CLIENTES
            AbrirFrmClientes (Index)
        
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
            Case 0: KEYBusqueda KeyAscii, 0 'cliente desde
            Case 1: KEYBusqueda KeyAscii, 1 'cliente hasta
        End Select
    Else
        KEYpress KeyAscii
    End If

End Sub

Private Sub KEYBusqueda(KeyAscii As Integer, indice As Integer)
    KeyAscii = 0
    imgBuscar_Click (indice)
End Sub


Private Sub txtCodigo_LostFocus(Index As Integer)
Dim cad As String, cadTipo As String 'tipo cliente

    'Quitar espacios en blanco por los lados
    txtCodigo(Index).Text = Trim(txtCodigo(Index).Text)
    
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub

    
    Select Case Index
            
        Case 0, 1 'clientes
            txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), "ssocio", "nomsocio", "codsocio", "N")
            If txtCodigo(Index).Text <> "" Then txtCodigo(Index).Text = Format(txtCodigo(Index).Text, "000000")
        
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
        .Opcion = Opcion
        .Show vbModal
    End With
End Sub

Private Sub AbrirFrmClientes(indice As Integer)
    indCodigo = indice
    Set frmCli = New frmManClien
    frmCli.DatosADevolverBusqueda = "0|1|"
    frmCli.DeConsulta = True
    frmCli.CodigoActual = txtCodigo(indCodigo)
    frmCli.Show vbModal
    Set frmCli = Nothing
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

Private Sub CargarParametros()
Dim i As Integer
Dim mes As Byte
Dim ano As Integer
Dim NomMes As String
Dim Fecha As Date

    mes = Combo1(0).ItemData(Combo1(0).ListIndex)
    ano = Combo1(1).Text
    For i = 5 To 1 Step -1
        If mes > 0 Then
            ano = ano
        Else
            mes = 12
            ano = CCur(ano) - 1
        End If
        
        Select Case i
            Case 5
                cadParam = cadParam & "mes5= " & mes & "|" & "ano5= " & ano & "|"
                numParam = numParam + 2
                mes5 = mes
                ano5 = ano
            Case 4
                cadParam = cadParam & "mes4= " & mes & "|" & "ano4= " & ano & "|"
                numParam = numParam + 2
                mes4 = mes
                ano4 = ano
            Case 3
                cadParam = cadParam & "mes3= " & mes & "|" & "ano3= " & ano & "|"
                numParam = numParam + 2
                mes3 = mes
                ano3 = ano
            Case 2
                cadParam = cadParam & "mes2= " & mes & "|" & "ano2= " & ano & "|"
                numParam = numParam + 2
                mes2 = mes
                ano2 = ano
            Case 1
                cadParam = cadParam & "mes1= " & mes & "|" & "ano1= " & ano & "|"
                numParam = numParam + 2
                mes1 = mes
                ano1 = ano
        End Select
        mes = mes - 1
    Next i
    
    desfec = "01/" & Format(mes1, "00") & "/" & Format(ano1, "0000")
    
    mes = mes5 + 1
    ano = ano5
    If mes > 12 Then
        mes = 1
        ano = ano + 1
    End If
    
    hasfec = "01/" & Format(mes, "00") & "/" & Format(ano, "0000")
    Fecha = CDate(hasfec) - 1
    hasfec = Format(Fecha, "dd/mm/yyyy")
    
    
    
    cadParam = cadParam & "pDHfechaFac= ""HASTA : " & Combo1(0).Text & " - " & Combo1(1).Text & """|"
    numParam = numParam + 1
    cadFormula = "{tmpinformes.codusu} = " & vSesion.Codigo
End Sub

Private Sub CargaCombo()
Dim cad As String
Dim RS As ADODB.Recordset
Dim i As Integer

    On Error GoTo ErrCarga
    Combo1(0).Clear
    
    'cargamos el combo del mes
    Combo1(0).AddItem "Enero"
    Combo1(0).ItemData(Combo1(0).NewIndex) = 1
    Combo1(0).AddItem "Febrero"
    Combo1(0).ItemData(Combo1(0).NewIndex) = 2
    Combo1(0).AddItem "Marzo"
    Combo1(0).ItemData(Combo1(0).NewIndex) = 3
    Combo1(0).AddItem "Abril"
    Combo1(0).ItemData(Combo1(0).NewIndex) = 4
    Combo1(0).AddItem "Mayo"
    Combo1(0).ItemData(Combo1(0).NewIndex) = 5
    Combo1(0).AddItem "Junio"
    Combo1(0).ItemData(Combo1(0).NewIndex) = 6
    Combo1(0).AddItem "Julio"
    Combo1(0).ItemData(Combo1(0).NewIndex) = 7
    Combo1(0).AddItem "Agosto"
    Combo1(0).ItemData(Combo1(0).NewIndex) = 8
    Combo1(0).AddItem "Septiembre"
    Combo1(0).ItemData(Combo1(0).NewIndex) = 9
    Combo1(0).AddItem "Octubre"
    Combo1(0).ItemData(Combo1(0).NewIndex) = 10
    Combo1(0).AddItem "Noviembre"
    Combo1(0).ItemData(Combo1(0).NewIndex) = 11
    Combo1(0).AddItem "Diciembre"
    Combo1(0).ItemData(Combo1(0).NewIndex) = 12
    
    
    Combo1(1).Clear
    'cargamos el combo del año
    For i = 0 To 50
        Combo1(1).AddItem Format(Year(Now) - i, "####")
        Combo1(1).ItemData(Combo1(1).NewIndex) = i
    Next i
    
    Exit Sub
    
ErrCarga:
    If Err.Number <> 0 Then MuestraError Err.Number, "Cargar datos combo.", Err.Description
End Sub

' el programa estaba pensado para que los datos se calcularan en el report al añadir el orden
' he tenido que cargar la temporal

Private Function CargarTemporal() As Boolean
Dim sql As String
Dim Sql2 As String
Dim sql3 As String
Dim Sql4 As String
Dim RS As ADODB.Recordset
Dim Rs2 As ADODB.Recordset
Dim Rs3 As ADODB.Recordset

Dim mes As Integer
Dim ano As Integer
Dim Fecha As Date
Dim AntSocio As Long
Dim ActSocio As Long
Dim Existe As String


    sql = "delete from tmpinformes where codusu = " & vSesion.Codigo
    Conn.Execute sql


    Pb1.visible = True
    CargarProgres Pb1, nRegs
    
    sql = "select codsocio, year(fecfactu), month(fecfactu), sum(totalfac) from schfac where fecfactu between " & DBSet(desfec, "F") & " and " & DBSet(hasfec, "F")
    If txtCodigo(0).Text <> "" Then sql = sql & " and codsocio >= " & DBSet(txtCodigo(0).Text, "N")
    If txtCodigo(1).Text <> "" Then sql = sql & " and codsocio <= " & DBSet(txtCodigo(1).Text, "N")
'    sql = sql & "(month(fecfactu) = " & mes5 & " and year(fecfactu) = " & ano5 & ") or "
'    sql = sql & "(month(fecfactu) = " & mes4 & " and year(fecfactu) = " & ano4 & ") or "
'    sql = sql & "(month(fecfactu) = " & mes3 & " and year(fecfactu) = " & ano3 & ") or "
'    sql = sql & "(month(fecfactu) = " & mes2 & " and year(fecfactu) = " & ano2 & ") or "
'    sql = sql & "(month(fecfactu) = " & mes1 & " and year(fecfactu) = " & ano1 & ") "
    sql = sql & " group by 1, 2, 3"
    sql = sql & " order by 1, 2, 3"
        
    Set RS = New ADODB.Recordset
    RS.Open sql, Conn, adOpenDynamic, adLockOptimistic, adCmdText
    If Not RS.EOF Then RS.MoveFirst
    AntSocio = RS.Fields(0).Value
    ActSocio = RS.Fields(0).Value
    
    While Not RS.EOF
        ActSocio = RS.Fields(0).Value
        If AntSocio <> ActSocio Then
            IncrementarProgres Pb1, 1
            AntSocio = ActSocio
        End If
        
        Existe = ""
        Existe = DevuelveDesdeBDNew(cPTours, "tmpinformes", "codigo1", "codusu", vSesion.Codigo, "N", , "codigo1", RS.Fields(0).Value, "N")
        If Existe = "" Then
            sql = "insert into tmpinformes (codusu,codigo1,importe6) values ("
            sql = sql & vSesion.Codigo & "," & DBSet(RS.Fields(0).Value, "N") & ",0)"
            Conn.Execute sql
        End If
        
        If RS.Fields(1).Value = ano1 And RS.Fields(2).Value = mes1 Then
           sql = "update tmpinformes set importe1 = " & DBSet(RS.Fields(3).Value, "N")
           sql = sql & " , importe6 = importe6 + " & DBSet(RS.Fields(3).Value, "N")
           sql = sql & " where codusu = " & vSesion.Codigo
           sql = sql & " and codigo1 = " & DBSet(RS.Fields(0).Value, "N")
        End If
        
        If RS.Fields(1).Value = ano2 And RS.Fields(2).Value = mes2 Then
           sql = "update tmpinformes set importe2 = " & DBSet(RS.Fields(3).Value, "N")
           sql = sql & " , importe6 = importe6 + " & DBSet(RS.Fields(3).Value, "N")
           sql = sql & " where codusu = " & vSesion.Codigo
           sql = sql & " and codigo1 = " & DBSet(RS.Fields(0).Value, "N")
        End If
        
        If RS.Fields(1).Value = ano3 And RS.Fields(2).Value = mes3 Then
           sql = "update tmpinformes set importe3 = " & DBSet(RS.Fields(3).Value, "N")
           sql = sql & " , importe6 = importe6 + " & DBSet(RS.Fields(3).Value, "N")
           sql = sql & " where codusu = " & vSesion.Codigo
           sql = sql & " and codigo1 = " & DBSet(RS.Fields(0).Value, "N")
        End If
        
        If RS.Fields(1).Value = ano4 And RS.Fields(2).Value = mes4 Then
           sql = "update tmpinformes set importe4 = " & DBSet(RS.Fields(3).Value, "N")
           sql = sql & " , importe6 = importe6 + " & DBSet(RS.Fields(3).Value, "N")
           sql = sql & " where codusu = " & vSesion.Codigo
           sql = sql & " and codigo1 = " & DBSet(RS.Fields(0).Value, "N")
        End If
        
        If RS.Fields(1).Value = ano5 And RS.Fields(2).Value = mes5 Then
           sql = "update tmpinformes set importe5 = " & DBSet(RS.Fields(3).Value, "N")
           sql = sql & " , importe6 = importe6 + " & DBSet(RS.Fields(3).Value, "N")
           sql = sql & " where codusu = " & vSesion.Codigo
           sql = sql & " and codigo1 = " & DBSet(RS.Fields(0).Value, "N")
        End If
        Conn.Execute sql
             
        RS.MoveNext
    
    Wend
    Pb1.visible = False
    Set RS = Nothing
    
End Function

