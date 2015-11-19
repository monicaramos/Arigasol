VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmDeshacerFact 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Deshacer Proceso de Facturación"
   ClientHeight    =   5850
   ClientLeft      =   45
   ClientTop       =   2430
   ClientWidth     =   6540
   Icon            =   "frmDeshacerFact.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5850
   ScaleWidth      =   6540
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
      Height          =   5670
      Left            =   90
      TabIndex        =   9
      Top             =   90
      Width           =   6375
      Begin VB.TextBox txtcodigo 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Index           =   6
         Left            =   2475
         MaxLength       =   50
         PasswordChar    =   "*"
         TabIndex        =   0
         Tag             =   "admon"
         Top             =   1485
         Width           =   1545
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   5
         Left            =   3720
         MaxLength       =   3
         TabIndex        =   2
         Top             =   2565
         Width           =   495
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   4
         Left            =   1575
         MaxLength       =   3
         TabIndex        =   1
         Top             =   2550
         Width           =   495
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   3
         Left            =   3735
         MaxLength       =   10
         TabIndex        =   6
         Tag             =   "Código Postal|T|S|||clientes|codposta|||"
         Top             =   3780
         Width           =   1050
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   2
         Left            =   1605
         MaxLength       =   10
         TabIndex        =   5
         Tag             =   "Código Postal|T|S|||clientes|codposta|||"
         Top             =   3750
         Width           =   1050
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   4755
         TabIndex        =   8
         Top             =   4980
         Width           =   975
      End
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   3570
         TabIndex        =   7
         Top             =   4980
         Width           =   975
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   0
         Left            =   1605
         MaxLength       =   7
         TabIndex        =   3
         Tag             =   "Nº de Factura|N|N|0|9999999|schfac|numfactu|0000000|S|"
         Top             =   3150
         Width           =   830
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   1
         Left            =   3720
         MaxLength       =   7
         TabIndex        =   4
         Tag             =   "Nº de Factura|N|N|0|9999999|schfac|numfactu|0000000|S|"
         Top             =   3180
         Width           =   830
      End
      Begin MSComctlLib.ProgressBar Pb1 
         Height          =   345
         Left            =   330
         TabIndex        =   19
         Top             =   4320
         Width           =   5595
         _ExtentX        =   9869
         _ExtentY        =   609
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Label Label3 
         Caption         =   "Password:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   345
         Left            =   1440
         TabIndex        =   22
         Top             =   1485
         Width           =   2235
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "No actualiza los contadores"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   510
         Left            =   405
         TabIndex        =   21
         Top             =   720
         Width           =   5595
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Este proceso borra facturas correlativas y las pasa a albaranes."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   465
         Left            =   360
         TabIndex        =   20
         Top             =   270
         Width           =   5640
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
         Left            =   330
         TabIndex        =   18
         Top             =   2310
         Width           =   1005
      End
      Begin VB.Label Label4 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   1
         Left            =   2880
         TabIndex        =   17
         Top             =   2565
         Width           =   420
      End
      Begin VB.Label Label4 
         Caption         =   "Desde"
         Height          =   195
         Index           =   0
         Left            =   720
         TabIndex        =   16
         Top             =   2550
         Width           =   465
      End
      Begin VB.Label Label4 
         Caption         =   "Fecha Factura"
         ForeColor       =   &H00972E0B&
         Height          =   255
         Index           =   16
         Left            =   330
         TabIndex        =   15
         Top             =   3510
         Width           =   1815
      End
      Begin VB.Label Label4 
         Caption         =   "Desde"
         Height          =   195
         Index           =   15
         Left            =   720
         TabIndex        =   14
         Top             =   3750
         Width           =   465
      End
      Begin VB.Label Label4 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   14
         Left            =   2880
         TabIndex        =   13
         Top             =   3780
         Width           =   420
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   2
         Left            =   1290
         Picture         =   "frmDeshacerFact.frx":000C
         ToolTipText     =   "Buscar fecha"
         Top             =   3750
         Width           =   240
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   3
         Left            =   3420
         Picture         =   "frmDeshacerFact.frx":0097
         ToolTipText     =   "Buscar fecha"
         Top             =   3780
         Width           =   240
      End
      Begin VB.Label Label4 
         Caption         =   "Desde"
         Height          =   195
         Index           =   13
         Left            =   720
         TabIndex        =   12
         Top             =   3150
         Width           =   465
      End
      Begin VB.Label Label4 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   12
         Left            =   2880
         TabIndex        =   11
         Top             =   3195
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
         Left            =   360
         TabIndex        =   10
         Top             =   2910
         Width           =   555
      End
   End
End
Attribute VB_Name = "frmDeshacerFact"
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
Dim I As Byte
Dim sql As String
Dim tipo As Byte
Dim NRegs As Long
Dim NumError As Long
Dim Mens As String

    If Not DatosOk Then Exit Sub

    cadSelect = ""

    'D/H Fecha factura
    cDesde = Trim(txtcodigo(2).Text)
    cHasta = Trim(txtcodigo(3).Text)
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{" & tabla & ".fecfactu}"
        TipCod = "F"
        If Not PonerDesdeHasta(cDesde, cHasta, "", "", "pDHfechaFac= """) Then Exit Sub
    End If
    
    'D/H letra de serie
    cDesde = Trim(txtcodigo(4).Text)
    cHasta = Trim(txtcodigo(5).Text)
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{" & tabla & ".letraser}"
        TipCod = "T"
        If Not PonerDesdeHasta(cDesde, cHasta, "", "", "pDHColec= """) Then Exit Sub
    End If
    
    'D/H numero de factura
    cDesde = Trim(txtcodigo(0).Text)
    cHasta = Trim(txtcodigo(1).Text)
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{" & tabla & ".numfactu}"
        TipCod = "N"
        If Not PonerDesdeHasta(cDesde, cHasta, "", "", "pDHColec= """) Then Exit Sub
    End If
    
    If Not HayRegParaInforme(tabla, cadSelect) Then Exit Sub
    
    Mens = "Deshacer Facturacion:"
    NumError = DeshacerFacturacion(Mens)
    
    
eError:
    If Err.Number <> 0 Or NumError <> 0 Then
        MuestraError Err.Number, "No se ha realizado el proceso de deshacer facturación. Llame a soporte." & vbCrLf & Mens
    Else
        MsgBox "Proceso realizado correctamente.", vbExclamation
    End If
    
    Pb1.visible = False
    cmdCancel_Click
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    If PrimeraVez Then
        PrimeraVez = False
        PonerFoco txtcodigo(6)
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
Dim h As Integer, w As Integer
Dim List As Collection

    PrimeraVez = True
    limpiar Me

    ActivarCLAVE
    
    'IMAGES para busqueda
'     Me.imgBuscar(8).Picture = frmPpal.imgListImages16.ListImages(1).Picture
     
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
    
    If txtcodigo(2).Text = "" Then
        txtcodigo(2).Text = Orden1 'fechaini del ejercicio de la conta
    End If
     
    If txtcodigo(3).Text = "" Then
        txtcodigo(3).Text = Orden2 'fecha fin del ejercicio de la conta
    End If
     

End Sub


Private Sub frmC_Selec(vFecha As Date)
 'Fecha
    txtcodigo(CByte(imgFec(2).Tag)).Text = Format(vFecha, "dd/MM/yyyy")
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
    If txtcodigo(Index).Text <> "" Then frmC.NovaData = txtcodigo(Index).Text

    frmC.Show vbModal
    Set frmC = Nothing
    PonerFoco txtcodigo(CByte(imgFec(2).Tag) + 2)
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

Private Sub Option1_Click(Index As Integer)
    Select Case Index
        Case 0
            Me.Caption = "Facturas por Cliente"
        Case 1
            Me.Caption = "Facturas por Tarjeta"
        Case 2
            Me.Caption = "Facturas por Cliente y por Tarjeta"
    End Select
    
End Sub

Private Sub txtCodigo_GotFocus(Index As Integer)
    ConseguirFoco txtcodigo(Index), 3
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
    txtcodigo(Index).Text = Trim(txtcodigo(Index).Text)
    
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub

    
    Select Case Index
        Case 2, 3  'FECHAS
            If txtcodigo(Index).Text <> "" Then PonerFormatoFecha txtcodigo(Index)
        
        Case 0, 1 ' NUMERO DE FACTURA
            If txtcodigo(Index).Text <> "" Then PonerFormatoEntero txtcodigo(Index)
        
        Case 4, 5 ' LETRA DE SERIE
            If txtcodigo(Index).Text <> "" Then txtcodigo(Index).Text = UCase(txtcodigo(Index).Text)
        
        Case 6
            If txtcodigo(Index).Text = "" Then Exit Sub
            If Trim(txtcodigo(Index).Text) <> Trim(txtcodigo(Index).Tag) Then
                MsgBox "    ACCESO DENEGADO    ", vbExclamation
                txtcodigo(Index).Text = ""
                PonerFoco txtcodigo(Index)
            Else
                DesactivarCLAVE
                PonerFoco txtcodigo(4)
            End If
        
    End Select
End Sub

Private Sub FrameCobrosVisible(visible As Boolean, ByRef h As Integer, ByRef w As Integer)
    Me.FrameCobros.visible = visible
    If visible = True Then
        Me.FrameCobros.Top = -90
        Me.FrameCobros.Left = 0
        Me.FrameCobros.Height = 5670
        Me.FrameCobros.Width = 6375
        w = Me.FrameCobros.Width
        h = Me.FrameCobros.Height
    End If
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



Private Function DeshacerFacturacion(ByRef Mens As String) As Long
Dim sql As String
Dim Sql2 As String
Dim Sql3 As String
Dim Sql4 As String
Dim Rs As ADODB.Recordset
Dim Socio As String
Dim Forpa As String

Dim AntLetraSer As String
Dim AntNumfactu As Long
Dim AntFecfactu As Date

Dim ActLetraSer As String
Dim ActNumfactu As Long
Dim ActFecfactu As Date

Dim NRegs As Long
Dim I As Long

    On Error GoTo eDeshacerFacturacion
    
    DeshacerFacturacion = 0
    Conn.BeginTrans

    sql = "select count(*) from slhfac where numalbar <> 'COOP.' "
    If txtcodigo(4).Text <> "" Then sql = sql & " and letraser >= " & DBSet(txtcodigo(4).Text, "T")
    If txtcodigo(5).Text <> "" Then sql = sql & " and letraser <= " & DBSet(txtcodigo(5).Text, "T")
    
    If txtcodigo(0).Text <> "" Then sql = sql & " and numfactu >= " & DBSet(txtcodigo(0).Text, "N")
    If txtcodigo(1).Text <> "" Then sql = sql & " and numfactu <= " & DBSet(txtcodigo(1).Text, "N")
    
    If txtcodigo(2).Text <> "" Then sql = sql & " and fecfactu >= " & DBSet(txtcodigo(2).Text, "F")
    If txtcodigo(3).Text <> "" Then sql = sql & " and fecfactu <= " & DBSet(txtcodigo(3).Text, "F")

    NRegs = TotalRegistros(sql)
    NRegs = NRegs + 1
    Pb1.Max = NRegs

    sql = "select * from slhfac where numalbar <> 'COOP.' "
    If txtcodigo(4).Text <> "" Then sql = sql & " and letraser >= " & DBSet(txtcodigo(4).Text, "T")
    If txtcodigo(5).Text <> "" Then sql = sql & " and letraser <= " & DBSet(txtcodigo(5).Text, "T")
    
    If txtcodigo(0).Text <> "" Then sql = sql & " and numfactu >= " & DBSet(txtcodigo(0).Text, "N")
    If txtcodigo(1).Text <> "" Then sql = sql & " and numfactu <= " & DBSet(txtcodigo(1).Text, "N")
    
    If txtcodigo(2).Text <> "" Then sql = sql & " and fecfactu >= " & DBSet(txtcodigo(2).Text, "F")
    If txtcodigo(3).Text <> "" Then sql = sql & " and fecfactu <= " & DBSet(txtcodigo(3).Text, "F")
    
    Set Rs = New ADODB.Recordset
    Rs.Open sql, Conn, adOpenStatic, adLockPessimistic, adCmdText
    
    Sql2 = "insert into scaalb (codclave, codsocio, numtarje, numalbar, fecalbar, horalbar, "
    Sql2 = Sql2 & "codturno, codartic, cantidad, preciove, importel, codforpa, matricul, "
    Sql2 = Sql2 & "codtraba, numfactu, numlinea, declaradogp) values "
    
    
    I = SugerirCodigoSiguienteStr("scaalb", "codclave")

    'contabilizar cada una de las facturas seleccionadas
    AntLetraSer = Rs!letraser
    AntNumfactu = Rs!numfactu
    AntFecfactu = Rs!fecfactu
    
    ActLetraSer = Rs!letraser
    ActNumfactu = Rs!numfactu
    ActFecfactu = Rs!fecfactu
    
    While Not Rs.EOF
        ActLetraSer = DBLet(Rs!letraser, "T")
        ActNumfactu = DBLet(Rs!numfactu, "N")
        ActFecfactu = DBLet(Rs!fecfactu, "F")
        
        If ActLetraSer <> AntLetraSer Or ActNumfactu <> AntNumfactu Or ActFecfactu <> AntFecfactu Then
            Sql4 = "delete from schfac where letraser = " & DBSet(AntLetraSer, "T") & " and "
            Sql4 = Sql4 & " numfactu = " & DBSet(AntNumfactu, "N") & " and "
            Sql4 = Sql4 & " fecfactu = " & DBSet(AntFecfactu, "F")
            
            Conn.Execute Sql4
            
            AntLetraSer = ActLetraSer
            AntNumfactu = ActNumfactu
            AntFecfactu = ActFecfactu
        End If
        
        If DBLet(Rs!numalbar, "T") <> "BONIFICA" Then
            Socio = ""
            Socio = DevuelveDesdeBDNew(cPTours, "schfac", "codsocio", "letraser", Rs!letraser, "T", , "numfactu", Rs!numfactu, "N", "fecfactu", Rs!fecfactu, "F")
            Forpa = ""
            Forpa = DevuelveDesdeBDNew(cPTours, "schfac", "codforpa", "letraser", Rs!letraser, "T", , "numfactu", Rs!numfactu, "N", "fecfactu", Rs!fecfactu, "F")
        
            Sql3 = DBSet(I, "N") & "," & DBSet(Socio, "N") & "," & DBSet(Rs!Numtarje, "N") & "," & DBSet(Rs!numalbar, "T") & ","
            Sql3 = Sql3 & DBSet(Rs!fecAlbar, "F") & "," & DBSet(Rs!horalbar, "FH") & "," & DBSet(Rs!codTurno, "N") & ","
            Sql3 = Sql3 & DBSet(Rs!codArtic, "N") & "," & DBSet(Rs!cantidad, "N") & "," & DBSet(Rs!preciove, "N") & ","
            Sql3 = Sql3 & DBSet(Rs!ImpLinea, "N") & "," & DBSet(Forpa, "N") & "," & ValorNulo & ","
            Sql3 = Sql3 & "0,0,0,1"
            
            Sql3 = "(" & Sql3 & "),"
            Sql2 = Sql2 & Sql3
            I = I + 1
        
            IncrementarProgres Me.Pb1, 1
        
        End If
        Sql4 = "delete from slhfac where letraser = " & DBSet(Rs!letraser, "T") & " and numfactu = "
        Sql4 = Sql4 & DBSet(Rs!numfactu, "N") & " and fecfactu = " & DBSet(Rs!fecfactu, "F")
        
        Conn.Execute Sql4
        
        Rs.MoveNext
    Wend
    
    'borramos la última cabecera
    Sql4 = "delete from schfac where letraser = " & DBSet(AntLetraSer, "T") & " and "
    Sql4 = Sql4 & " numfactu = " & DBSet(AntNumfactu, "N") & " and "
    Sql4 = Sql4 & " fecfactu = " & DBSet(AntFecfactu, "F")
    
    Conn.Execute Sql4
    
    
    ' quitamos la ultima coma de la ejecucion
    Sql2 = Mid(Sql2, 1, Len(Sql2) - 1)
    Conn.Execute Sql2
    
    Conn.CommitTrans
    Exit Function

eDeshacerFacturacion:
    If Err.Number <> 0 Then
         DeshacerFacturacion = Err.Number
         Mens = Err.Description
         Conn.RollbackTrans
    End If
End Function

Private Sub ActivarCLAVE()
Dim I As Integer
    
    For I = 0 To 5
        txtcodigo(I).Enabled = False
    Next I
    txtcodigo(6).Enabled = True
    
    cmdAceptar.Enabled = False
    cmdCancel.Enabled = True

End Sub

Private Sub DesactivarCLAVE()
Dim I As Integer

    For I = 0 To 5
        txtcodigo(I).Enabled = True
    Next I
    txtcodigo(6).Enabled = False
    
    cmdAceptar.Enabled = True
End Sub


Private Function DatosOk() As Boolean
    DatosOk = False
    If (txtcodigo(4).Text <> "" And txtcodigo(5) <> "") And _
       (txtcodigo(0).Text <> "" And txtcodigo(1) <> "") And _
       (txtcodigo(2).Text <> "" And txtcodigo(3) <> "") Then
       DatosOk = True
    Else
        MsgBox "Debe introducir obligatoriamente todos los rangos de valores. Reintroduzca.", vbExclamation
        PonerFoco txtcodigo(4)
    End If
End Function
