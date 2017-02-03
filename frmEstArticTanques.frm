VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmEstArticTanques 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Estad�stica de Art�culos en Tanques"
   ClientHeight    =   4545
   ClientLeft      =   45
   ClientTop       =   2430
   ClientWidth     =   7125
   Icon            =   "frmEstArticTanques.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4545
   ScaleWidth      =   7125
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
      TabIndex        =   6
      Top             =   120
      Width           =   6915
      Begin VB.CheckBox ChkResumen 
         Caption         =   "Resumen por Fechas"
         Height          =   285
         Left            =   510
         TabIndex        =   16
         Top             =   2880
         Width           =   2685
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   1
         Left            =   2865
         Locked          =   -1  'True
         TabIndex        =   12
         Text            =   "Text5"
         Top             =   2430
         Width           =   3555
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   0
         Left            =   2865
         Locked          =   -1  'True
         TabIndex        =   11
         Text            =   "Text5"
         Top             =   2070
         Width           =   3585
      End
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Index           =   1
         Left            =   1740
         MaxLength       =   16
         TabIndex        =   3
         Tag             =   "C�digo de articulo|N|N|0|999999|sartic|codartic|000000|S|"
         Top             =   2430
         Width           =   1065
      End
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Index           =   0
         Left            =   1740
         MaxLength       =   16
         TabIndex        =   2
         Tag             =   "C�digo de articulo|N|N|0|999999|sartic|codartic|000000|S|"
         Top             =   2070
         Width           =   1065
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   3
         Left            =   1755
         MaxLength       =   10
         TabIndex        =   1
         Tag             =   "C�digo Postal|T|S|||clientes|codposta|||"
         Top             =   1440
         Width           =   1050
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   2
         Left            =   1755
         MaxLength       =   10
         TabIndex        =   0
         Tag             =   "C�digo Postal|T|S|||clientes|codposta|||"
         Top             =   1080
         Width           =   1050
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   5415
         TabIndex        =   5
         Top             =   3660
         Width           =   975
      End
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   4230
         TabIndex        =   4
         Top             =   3660
         Width           =   975
      End
      Begin MSComctlLib.ProgressBar Pb1 
         Height          =   255
         Left            =   480
         TabIndex        =   10
         Top             =   3330
         Width           =   5970
         _ExtentX        =   10530
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   1
         Left            =   1440
         Top             =   2460
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   0
         Left            =   1440
         Top             =   2070
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Art�culo"
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
         Left            =   480
         TabIndex        =   15
         Top             =   1830
         Width           =   540
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   22
         Left            =   870
         TabIndex        =   14
         Top             =   2430
         Width           =   420
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
         Height          =   195
         Index           =   23
         Left            =   870
         TabIndex        =   13
         Top             =   2070
         Width           =   465
      End
      Begin VB.Label Label4 
         Caption         =   "Fecha"
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
         Height          =   255
         Index           =   16
         Left            =   480
         TabIndex        =   9
         Top             =   660
         Width           =   1815
      End
      Begin VB.Label Label4 
         Caption         =   "Desde"
         Height          =   195
         Index           =   15
         Left            =   870
         TabIndex        =   8
         Top             =   1080
         Width           =   465
      End
      Begin VB.Label Label4 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   14
         Left            =   870
         TabIndex        =   7
         Top             =   1440
         Width           =   420
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   2
         Left            =   1440
         Picture         =   "frmEstArticTanques.frx":000C
         ToolTipText     =   "Buscar fecha"
         Top             =   1080
         Width           =   240
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   3
         Left            =   1440
         Picture         =   "frmEstArticTanques.frx":0097
         ToolTipText     =   "Buscar fecha"
         Top             =   1440
         Width           =   240
      End
   End
End
Attribute VB_Name = "frmEstArticTanques"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' +-+-+-+-+-+-+-+-+-+-+-
' +-+- Autor MANOLO +-+-
' +-+-+-+-+-+-+-+-+-+-+-

Option Explicit

Public NumCod As String 'Para indicar cod. Traspaso,Movimiento, etc. que llama
                        'Para indicar n� oferta a imprimir

Public CadTag As String 'Cadena con el Tag del campo que se va a poner en D/H en los listados
                        'Se necesita si el tipo de codigo es texto

    
Private Conexion As Byte
'1.- Conexi�n a BD Ariges  2.- Conexi�n a BD Conta

Private HaDevueltoDatos As Boolean

Private WithEvents frmArt As frmManArtic 'Articulos
Attribute frmArt.VB_VarHelpID = -1
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
Dim indFrame As Single 'n� de frame en el que estamos
 
'Se inicializan para cada Informe (tabla de BD a la que hace referencia
Dim tabla As String
Dim Codigo As String 'C�digo para FormulaSelection de Crystal Report
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
Dim SQL As String

    InicializarVbles
    
    '========= PARAMETROS  =============================
    'A�adir el parametro de Empresa
    cadParam = cadParam & "|pEmpresa=""" & vEmpresa.nomEmpre & """|"
    numParam = numParam + 1
    
    'D/H Fecha albaran
    cDesde = Trim(txtCodigo(2).Text)
    cHasta = Trim(txtCodigo(3).Text)
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{sturno.fechatur}"
        TipCod = "F"
        If Not PonerDesdeHasta(cDesde, cHasta, "", "", "pDHFecha= """) Then Exit Sub
    End If

    'D/H Art�culo
    cDesde = Trim(txtCodigo(0).Text)
    cHasta = Trim(txtCodigo(1).Text)
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{sturno.codartic}"
        TipCod = "N"
        If Not PonerDesdeHasta(cDesde, cHasta, "", "", "pDHArtic= """) Then Exit Sub
    End If

    If HayRegParaInforme("sturno", cadSelect) Then
        cadParam = cadParam & "pResumen=" & Me.ChkResumen.Value & "|"
        numParam = numParam + 1
    
    
        cadTitulo = "Estad�stica de Art�culos de Tanques"
        cadNombreRPT = "rEstArticulosTurnos.rpt"
        LlamarImprimir
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    If PrimeraVez Then
        PrimeraVez = False
        PonerFoco txtCodigo(2)
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

    '###Descomentar
'    CommitConexion
         
    FrameCobrosVisible True, h, w
    indFrame = 5
    
    Pb1.visible = False
    
    'Esto se consigue poneinedo el cancel en el opcion k corresponda
    Me.cmdCancel.Cancel = True
'    Me.Width = w + 70
'    Me.Height = h + 350
End Sub

Private Sub frmArt_DatoSeleccionado(CadenaSeleccion As String)
'Form de Consulta de Familias
    txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000000")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmC_Selec(vFecha As Date)
 'Fecha
    txtCodigo(CByte(imgFec(2).Tag)).Text = Format(vFecha, "dd/MM/yyyy")
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
        
        Case 0, 1 'ARTICULOS
            AbrirFrmArticulos (Index)
        
    End Select
    PonerFoco txtCodigo(indCodigo)
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
            Case 0: KEYBusqueda KeyAscii, 0 'articulo desde
            Case 1: KEYBusqueda KeyAscii, 1 'articulo hasta
            Case 2: KEYFecha KeyAscii, 2 'fecha factura desde
            Case 3: KEYFecha KeyAscii, 3 'fecha factura hasta
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
        Case 0, 1 'ARTICULOS
            txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), "sartic", "nomartic", "codartic", "N")
            If txtCodigo(Index).Text <> "" Then txtCodigo(Index).Text = Format(txtCodigo(Index).Text, "000000")
    
    
        Case 2, 3 'FECHAS
            If txtCodigo(Index).Text <> "" Then PonerFormatoFecha txtCodigo(Index)
            
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
'A�ade a cadFormula y cadSelect la cadena de seleccion:
'       "(codigo>=codD AND codigo<=codH)"
' y a�ade a cadParam la cadena para mostrar en la cabecera informe:
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
        .Opcion = 1
        .Show vbModal
    End With
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


Private Function CargarTemporal() As Boolean

Dim SQL As String
Dim sql2 As String
Dim Sql3 As String
Dim Sql4 As String
Dim Rs As ADODB.Recordset
Dim Rs2 As ADODB.Recordset
Dim Rs3 As ADODB.Recordset
Dim vImp1 As Currency
Dim vImpb1 As Currency
Dim vImp2 As Currency
Dim vImpb2 As Currency
Dim vImp3 As Currency
Dim vImpb3 As Currency
Dim vImp4 As Currency
Dim vImpb4 As Currency
Dim vImp5 As Currency
Dim vImpb5 As Currency
Dim vImp6 As Currency
Dim NRegs As Integer

    SQL = "delete from tmpinformes where codusu = " & vSesion.Codigo
    Conn.Execute SQL

    SQL = "select count(*) from sartic where codfamia in (select codfamia from sfamia where tipfamia = 1) "
    NRegs = TotalRegistros(SQL)
    
    Pb1.visible = True
    CargarProgres Pb1, NRegs
        
    SQL = "select codartic, nomartic from sartic where codfamia in (select codfamia from sfamia where tipfamia = 1) "
    
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    While Not Rs.EOF
        IncrementarProgres Pb1, 1
        
        Sql3 = "SELECT ssocio.grupoestartic, sum(cantidad), sum(implinea) "
        Sql3 = Sql3 & " from slhfac, schfac, ssocio "
        Sql3 = Sql3 & " where slhfac.letraser = schfac.letraser and slhfac.numfactu = schfac.numfactu "
        Sql3 = Sql3 & " and slhfac.fecfactu = schfac.fecfactu and slhfac.codartic = " & DBSet(Rs!codArtic, "N")
        
        If txtCodigo(2).Text <> "" Then Sql3 = Sql3 & " and slhfac.fecalbar >= " & DBSet(txtCodigo(2).Text, "F")
        If txtCodigo(3).Text <> "" Then Sql3 = Sql3 & " and slhfac.fecalbar <= " & DBSet(txtCodigo(3).Text, "F")
        
        Sql3 = Sql3 & " and schfac.codsocio = ssocio.codsocio "
        Sql3 = Sql3 & " group by  ssocio.grupoestartic"
        Sql3 = Sql3 & " order by ssocio.grupoestartic "
 
        Set Rs2 = New ADODB.Recordset
        Rs2.Open Sql3, Conn, adOpenDynamic, adLockOptimistic, adCmdText
       
        vImp1 = 0
        vImpb1 = 0
        vImp2 = 0
        vImpb2 = 0
        vImp3 = 0
        vImpb3 = 0
        vImp4 = 0
        vImpb4 = 0
        vImp5 = 0
        vImpb5 = 0

        
        While Not Rs2.EOF
            Select Case Rs2!grupoestartic
                Case 0
                    vImp1 = Rs2.Fields(1).Value
                    vImpb1 = Rs2.Fields(2).Value
                Case 1
                    vImp2 = Rs2.Fields(1).Value
                    vImpb2 = Rs2.Fields(2).Value
                Case 2
                    vImp3 = Rs2.Fields(1).Value
                    vImpb3 = Rs2.Fields(2).Value
                Case 3
                    vImp4 = Rs2.Fields(1).Value
                    vImpb4 = Rs2.Fields(2).Value
                Case 4
                    vImp5 = Rs2.Fields(1).Value
                    vImpb5 = Rs2.Fields(2).Value
            End Select
            
            Rs2.MoveNext
        Wend
        
        Set Rs2 = Nothing
        
        Sql4 = "select sum(litrosve) from sturno where codartic = " & DBSet(Rs!codArtic, "N") & " and tiporegi = 3 "
        If txtCodigo(2).Text <> "" Then Sql4 = Sql4 & " and fechatur >= " & DBSet(txtCodigo(2).Text, "F")
        If txtCodigo(3).Text <> "" Then Sql4 = Sql4 & " and fechatur <= " & DBSet(txtCodigo(3).Text, "F")
        Set Rs3 = New ADODB.Recordset
        Rs3.Open Sql4, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        If Rs3.EOF Then
            vImp6 = 0
        Else
            vImp6 = DBLet(Rs3.Fields(0).Value, "N")
        End If
        Set Rs3 = Nothing
        
        
        sql2 = "insert into tmpinformes (codusu, codigo1, nombre1, importe1, importe2, importe3, importe4, importe5, importe6, "
        sql2 = sql2 & "importeb1, importeb2, importeb3, importeb4, importeb5) values (" & vSesion.Codigo & ","
        sql2 = sql2 & DBSet(Rs!codArtic, "N") & "," & DBSet(Rs!NomArtic, "T") & ","
        sql2 = sql2 & DBSet(vImp1, "N") & "," & DBSet(vImp2, "N") & "," & DBSet(vImp3, "N") & "," & DBSet(vImp4, "N") & ","
        sql2 = sql2 & DBSet(vImp5, "N") & "," & DBSet(vImp6, "N") & ","
        sql2 = sql2 & DBSet(vImpb1, "N") & "," & DBSet(vImpb2, "N") & "," & DBSet(vImpb3, "N") & "," & DBSet(vImpb4, "N") & ","
        sql2 = sql2 & DBSet(vImpb5, "N") & ")"
        
        Conn.Execute sql2
    
        Rs.MoveNext
    
    Wend
    Pb1.visible = False
    Set Rs = Nothing
    
End Function
