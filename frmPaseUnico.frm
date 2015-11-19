VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmPaseUnico 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Traspaso a Unicoo"
   ClientHeight    =   4530
   ClientLeft      =   45
   ClientTop       =   2430
   ClientWidth     =   6645
   Icon            =   "frmPaseUnico.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4530
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
      Height          =   4245
      Left            =   90
      TabIndex        =   8
      Top             =   120
      Width           =   6375
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   5
         Left            =   3690
         MaxLength       =   3
         TabIndex        =   1
         Top             =   795
         Width           =   495
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   4
         Left            =   1575
         MaxLength       =   3
         TabIndex        =   0
         Top             =   780
         Width           =   495
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   3
         Left            =   3690
         MaxLength       =   10
         TabIndex        =   5
         Tag             =   "Código Postal|T|S|||clientes|codposta|||"
         Top             =   2010
         Width           =   1050
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   2
         Left            =   1575
         MaxLength       =   10
         TabIndex        =   4
         Tag             =   "Código Postal|T|S|||clientes|codposta|||"
         Top             =   1980
         Width           =   1050
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   4695
         TabIndex        =   7
         Top             =   3540
         Width           =   975
      End
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   3510
         TabIndex        =   6
         Top             =   3540
         Width           =   975
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   0
         Left            =   1575
         MaxLength       =   7
         TabIndex        =   2
         Tag             =   "Nº de Factura|N|N|0|9999999|schfac|numfactu|0000000|S|"
         Top             =   1380
         Width           =   830
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   1
         Left            =   3690
         MaxLength       =   7
         TabIndex        =   3
         Tag             =   "Nº de Factura|N|N|0|9999999|schfac|numfactu|0000000|S|"
         Top             =   1410
         Width           =   830
      End
      Begin MSComctlLib.ProgressBar Pb1 
         Height          =   255
         Left            =   270
         TabIndex        =   18
         Top             =   2580
         Width           =   5430
         _ExtentX        =   9578
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   300
         Top             =   3120
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Label lblProgres 
         Height          =   285
         Index           =   1
         Left            =   300
         TabIndex        =   20
         Top             =   3150
         Width           =   5295
      End
      Begin VB.Label lblProgres 
         Height          =   285
         Index           =   0
         Left            =   300
         TabIndex        =   19
         Top             =   2880
         Width           =   5265
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Letra de Serie"
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
         TabIndex        =   17
         Top             =   540
         Width           =   1005
      End
      Begin VB.Label Label4 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   1
         Left            =   2850
         TabIndex        =   16
         Top             =   795
         Width           =   420
      End
      Begin VB.Label Label4 
         Caption         =   "Desde"
         Height          =   195
         Index           =   0
         Left            =   690
         TabIndex        =   15
         Top             =   780
         Width           =   465
      End
      Begin VB.Label Label4 
         Caption         =   "Fecha Factura"
         ForeColor       =   &H00972E0B&
         Height          =   255
         Index           =   16
         Left            =   300
         TabIndex        =   14
         Top             =   1740
         Width           =   1815
      End
      Begin VB.Label Label4 
         Caption         =   "Desde"
         Height          =   195
         Index           =   15
         Left            =   690
         TabIndex        =   13
         Top             =   1980
         Width           =   465
      End
      Begin VB.Label Label4 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   14
         Left            =   2850
         TabIndex        =   12
         Top             =   2010
         Width           =   420
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   2
         Left            =   1260
         Picture         =   "frmPaseUnico.frx":000C
         ToolTipText     =   "Buscar fecha"
         Top             =   1980
         Width           =   240
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   3
         Left            =   3390
         Picture         =   "frmPaseUnico.frx":0097
         ToolTipText     =   "Buscar fecha"
         Top             =   2010
         Width           =   240
      End
      Begin VB.Label Label4 
         Caption         =   "Desde"
         Height          =   195
         Index           =   13
         Left            =   690
         TabIndex        =   11
         Top             =   1380
         Width           =   465
      End
      Begin VB.Label Label4 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   12
         Left            =   2850
         TabIndex        =   10
         Top             =   1425
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
         Index           =   11
         Left            =   300
         TabIndex        =   9
         Top             =   1140
         Width           =   555
      End
   End
End
Attribute VB_Name = "frmPaseUnico"
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
Private WithEvents frmcli As frmManClien 'Clientes
Attribute frmcli.VB_VarHelpID = -1
Private WithEvents frmC As frmCal 'calendario fechas
Attribute frmC.VB_VarHelpID = -1
Private WithEvents frmBpr As frmManBanco 'Banco Propio
Attribute frmBpr.VB_VarHelpID = -1

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
Dim cadMen As String
Dim i As Byte
Dim Sql As String
Dim Tipo As Byte
Dim nRegs As Long
Dim NumError As Long

    If Not DatosOk Then Exit Sub
    
    cadSelect = tabla & ".intconta=0 "
    
    'D/H Fecha factura
    cDesde = Trim(txtCodigo(2).Text)
    cHasta = Trim(txtCodigo(3).Text)
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{" & tabla & ".fecfactu}"
        TipCod = "F"
        If Not PonerDesdeHasta(cDesde, cHasta, "", "", "pDHfechaFac= """) Then Exit Sub
    End If
    
    'D/H letra de serie
    cDesde = Trim(txtCodigo(4).Text)
    cHasta = Trim(txtCodigo(5).Text)
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{" & tabla & ".letraser}"
        TipCod = "T"
        If Not PonerDesdeHasta(cDesde, cHasta, "", "", "pDHColec= """) Then Exit Sub
    End If
    
    'D/H numero de factura
    cDesde = Trim(txtCodigo(0).Text)
    cHasta = Trim(txtCodigo(1).Text)
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{" & tabla & ".numfactu}"
        TipCod = "N"
        If Not PonerDesdeHasta(cDesde, cHasta, "", "", "pDHColec= """) Then Exit Sub
    End If
    
    If Not HayRegParaInforme(tabla, cadSelect) Then Exit Sub
    
    If Not ExistenCodigosExternos(tabla, cadSelect) Then Exit Sub
    
    
    If GeneraFichero(tabla, cadSelect) Then
        If CopiarFichero Then
            If ActualizarRegistros(tabla, cadSelect) Then
                MsgBox "Proceso realizado correctamente", vbExclamation
                cmdCancel_Click
            End If
        End If
    End If
    'Desbloqueamos ya no estamos contabilizando facturas
    DesBloqueoManual ("VENCON") 'VENtas CONtabilizar
    
eError:
    If Err.Number <> 0 Or NumError <> 0 Then
        MuestraError Err.Number, "No se ha realizado el proceso de contabilización. Llame a soporte."
    End If
    
    Pb1.visible = False
    lblProgres(0).Caption = ""
    lblProgres(1).Caption = ""
    cmdCancel_Click
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
    Limpiar Me

    '###Descomentar
'    CommitConexion
         
    FrameCobrosVisible True, h, w
    indFrame = 5
    tabla = "schfac"
    
    Pb1.visible = False
    
    
    'Esto se consigue poneinedo el cancel en el opcion k corresponda
    Me.cmdCancel.Cancel = True
    Me.Width = w + 70
    Me.Height = h + 350
    
     
     If txtCodigo(2).Text = "" Then
        txtCodigo(2).Text = Orden1 'fechaini del ejercicio de la conta
     End If
     
     If txtCodigo(3).Text = "" Then
        txtCodigo(3).Text = Orden2 'fecha fin del ejercicio de la conta
     End If
     

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
'14/02/2007 antes estaba esto
'    KEYpress KeyAscii
' ahora
    If KeyAscii = teclaBuscar Then
        Select Case Index
            Case 2: KEYFecha KeyAscii, 2 'fecha desde
            Case 3: KEYFecha KeyAscii, 3 'fecha hasta
        End Select
    Else
        KEYpress KeyAscii
    End If

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
        
        Case 0, 1 ' NUMERO DE FACTURA
            If txtCodigo(Index).Text <> "" Then PonerFormatoEntero txtCodigo(Index)
        
        Case 4, 5 ' LETRA DE SERIE
            If txtCodigo(Index).Text <> "" Then txtCodigo(Index).Text = UCase(txtCodigo(Index).Text)
        
        
    End Select
End Sub

Private Sub FrameCobrosVisible(visible As Boolean, ByRef h As Integer, ByRef w As Integer)
    Me.FrameCobros.visible = visible
    If visible = True Then
        Me.FrameCobros.Top = -90
        Me.FrameCobros.Left = 0
        Me.FrameCobros.Height = 4245
        Me.FrameCobros.Width = 6375
        w = Me.FrameCobros.Width
        h = Me.FrameCobros.Height
    End If
End Sub

Private Sub ValoresPorDefecto()
'    txtCodigo(7).Text = Format(Now, "dd/mm/yyyy")
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

Private Sub AbrirFrmBancoPropio(indice As Integer)
    indCodigo = indice
    Set frmBpr = New frmManBanco
    frmBpr.DatosADevolverBusqueda = "0|1|"
    frmBpr.DeConsulta = True
    frmBpr.CodigoActual = txtCodigo(indCodigo)
    frmBpr.Show vbModal
    Set frmBpr = Nothing
End Sub

Private Function DatosOk() As Boolean
Dim b As Boolean

    b = True

    '07022007 he añadido esto tambien aquí
     If txtCodigo(2).Text = "" Then
        txtCodigo(2).Text = Orden1 'fechaini del ejercicio de la conta
     End If
     
     If txtCodigo(3).Text = "" Then
        txtCodigo(3).Text = Orden2 'fecha fin del ejercicio de la conta
     End If

    DatosOk = b
End Function

Private Function GeneraFichero(vTabla As String, vSelect As String) As Boolean
Dim NFich1 As Integer
Dim NFich2 As Integer
Dim Rs As ADODB.Recordset
Dim cad As String
Dim Sql As String
Dim AntLetraSer As String
Dim ActLetraSer As String
Dim AntNumfactu As Long
Dim ActNumfactu As Long
Dim v_Hayreg As Integer
Dim AntTarjet As Long
Dim ActTarjet As Long
Dim AntFecfactu As Date
Dim ActFecfactu As Date

Dim vsocio As String

Dim NomSocio As String
Dim NomArtic As String
Dim b As Boolean
Dim Mens As String

    On Error GoTo EGen
    
    GeneraFichero = False

    NFich1 = FreeFile
    Open App.path & "\traspaso.csv" For Output As #NFich1
    

    Set Rs = New ADODB.Recordset
    
    'partimos de la tabla de historico de facturas
    '[Monica]12/02/2015: agrupamos por factura y codigo externo del articulo
    'antes
'    sql = "SELECT slhfac.letraser, slhfac.numfactu, slhfac.fecfactu, slhfac.codartic, "
'    sql = sql & " slhfac.cantidad, slhfac.preciove, slhfac.implinea, "
'    sql = sql & " schfac.codsocio, schfac.codforpa, sforpa.codexterno as fpexterno, ssocio.codexterno as sexterno, sartic.codexterno as aexterno "
'    sql = sql & " slhfac.cantidad, slhfac.preciove, slhfac.implinea, "

    'ahora
    Sql = "SELECT slhfac.letraser, slhfac.numfactu, slhfac.fecfactu, sartic.codexterno as aexterno, sforpa.codexterno as fpexterno, ssocio.codexterno as sexterno,  "
    Sql = Sql & " sum(slhfac.cantidad) cantidad, sum(slhfac.implinea) implinea "
    
    Sql = Sql & " from (((" & vTabla & " inner join slhfac on schfac.letraser = slhfac.letraser and schfac.numfactu = slhfac.numfactu and schfac.fecfactu = slhfac.fecfactu)"
    Sql = Sql & " INNER JOIN ssocio ON ssocio.codsocio = schfac.codsocio) "
    Sql = Sql & " INNER JOIN sforpa ON sforpa.codforpa = schfac.codforpa) "
    Sql = Sql & " INNER JOIN sartic ON slhfac.codartic = sartic.codartic"
    Sql = Sql & " where " & vSelect
    Sql = Sql & " group by 1,2,3,4,5,6"
    Sql = Sql & " order by 1,2,3,4 "

    Rs.Open Sql, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    
    Dim nRegs As Integer
    
    nRegs = TotalRegistrosConsulta(Sql)
    Me.Pb1.visible = True
    CargarProgres Me.Pb1, nRegs
    
    AntLetraSer = ""
    ActLetraSer = DBLet(Rs!letraser, "T")
    AntNumfactu = 0
    ActNumfactu = DBLet(Rs!numfactu, "N")
    AntFecfactu = "1900-01-01"
    ActFecfactu = DBLet(Rs!fecfactu, "F")
    
    b = True
    v_Hayreg = 0
    While Not Rs.EOF And b
        v_Hayreg = 1
        
        ActLetraSer = DBLet(Rs!letraser)
        ActNumfactu = DBLet(Rs!numfactu)
        ActFecfactu = DBLet(Rs!fecfactu, "F")
        
        IncrementarProgres Pb1, 1
        
        cad = ""
        If AntLetraSer <> ActLetraSer Or AntNumfactu <> ActNumfactu Or AntFecfactu <> ActFecfactu Then
            Dim FecFact As String
            
            FecFact = Mid(Rs!fecfactu, 7, 4) & Mid(Rs!fecfactu, 4, 2) & Mid(Rs!fecfactu, 1, 2)
            
            cad = "V;106;" & ActLetraSer & ";" & ActLetraSer & "-" & ActNumfactu & ";" & DBLet(Rs!fpexterno) & ";" & DBLet(Rs!sexterno) & ";"
            cad = cad & FecFact & ";2;"
            
            Print #NFich1, cad
            
            AntLetraSer = ActLetraSer
            AntNumfactu = ActNumfactu
            AntFecfactu = ActFecfactu
        Else
'            cad = ";;;;;;;;"
        End If
    
        cad = "D;" & DBLet(Rs!aexterno) & ";1,00;"
        'cad = cad & Replace(DBLet(RS!cantidad), ".", ",") & ";"
        cad = cad & Replace(Format(DBLet(Rs!ImpLinea), "#######0.00"), ".", ",") & ";"
    
        Print #NFich1, cad
            
        Rs.MoveNext
    Wend
    
    
EGen:
    Close (NFich1)
    Set Rs = Nothing
    Me.Pb1.visible = False
    If Err.Number <> 0 Then
        MuestraError Err.Number, Err.Description & vbCrLf & Mens
    Else
        GeneraFichero = True
    End If

End Function


Public Function CopiarFichero() As Boolean
Dim nomfich As String
Dim cadena As String
On Error GoTo ecopiarfichero

    CopiarFichero = False
    ' abrimos el commondialog para indicar donde guardarlo
'    Me.CommonDialog1.InitDir = App.path

    Me.CommonDialog1.DefaultExt = "csv"
    cadena = Format(txtCodigo(2).Text, FormatoFecha)
    CommonDialog1.Filter = "Archivos csv|csv|"
    CommonDialog1.FilterIndex = 1
    
    ' copiamos el primer fichero
    CommonDialog1.FileName = "traspaso.csv"
    
    Me.CommonDialog1.ShowSave
    If CommonDialog1.FileName <> "" Then
        FileCopy App.path & "\traspaso.csv", CommonDialog1.FileName
    End If
    
    CopiarFichero = True


ecopiarfichero:
    If Err.Number <> 0 Then
        MuestraError Err.Number, Err.Description
    End If
    Err.Clear

End Function

Private Function ActualizarRegistros(vTabla As String, vSelect As String) As Boolean
Dim Rs As ADODB.Recordset
Dim cad As String
Dim Sql As String

    On Error GoTo EGen
    
    ActualizarRegistros = False


    Sql = "update " & vTabla
    Sql = Sql & " set intconta = 1 "
    Sql = Sql & " where " & vSelect
    
    Conn.Execute Sql

EGen:
    If Err.Number <> 0 Then
        MuestraError Err.Number, Err.Description
    Else
        ActualizarRegistros = True
    End If
End Function


Private Function ExistenCodigosExternos(vTabla As String, vSelect As String) As Boolean
Dim Rs As ADODB.Recordset
Dim cad As String
Dim Sql As String
Dim CadAux As String
Dim Mens As String


    On Error GoTo EGen
    
    ExistenCodigosExternos = False


    ' Socios
    Sql = "select codsocio from ssocio where codsocio in (select codsocio from schfac where " & vSelect & ") and (codexterno is null or codexterno = '')"
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    CadAux = ""
    While Not Rs.EOF
        CadAux = CadAux & DBSet(Rs!codsocio, "N") & ", "
        
        Rs.MoveNext
    Wend
    Set Rs = Nothing
    
    If CadAux <> "" Then
        Mens = "Los socios siguientes no tienen código externo. Revise. " & vbCrLf & vbCrLf & Mid(CadAux, 1, Len(CadAux) - 2)
        
        MsgBox Mens, vbExclamation
        Exit Function
    End If
    
    
    ' Formas de Pago
    Sql = "select codforpa from sforpa where codforpa in (select codforpa from schfac where " & vSelect & ") and (codexterno is null or codexterno = '')"
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    CadAux = ""
    While Not Rs.EOF
        CadAux = CadAux & DBSet(Rs!Codforpa, "N") & ", "
        
        Rs.MoveNext
    Wend
    Set Rs = Nothing
    
    If CadAux <> "" Then
        Mens = "Las formas de pago siguientes no tienen código externo. Revise. " & vbCrLf & vbCrLf & Mid(CadAux, 1, Len(CadAux) - 2)
        
        MsgBox Mens, vbExclamation
        Exit Function
    End If
    
    ' Articulos
    Sql = "select codartic from sartic where codartic in (select codartic from slhfac inner join schfac on slhfac.letraser = schfac.letraser and slhfac.numfactu = schfac.numfactu and slhfac.fecfactu = schfac.fecfactu where " & vSelect & ") and (codexterno is null or codexterno = '')"
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    CadAux = ""
    While Not Rs.EOF
        CadAux = CadAux & DBSet(Rs!codArtic, "N") & ", "
        
        Rs.MoveNext
    Wend
    Set Rs = Nothing
    
    If CadAux <> "" Then
        Mens = "Los artículos siguientes no tienen código externo. Revise. " & vbCrLf & vbCrLf & Mid(CadAux, 1, Len(CadAux) - 2)
        
        MsgBox Mens, vbExclamation
        Exit Function
    End If
    
    
EGen:
    If Err.Number <> 0 Then
        MuestraError Err.Number, Err.Description & vbCrLf & Mens
    Else
        ExistenCodigosExternos = True
    End If
End Function


