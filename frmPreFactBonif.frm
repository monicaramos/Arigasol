VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmPreFactBonif 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Prefacturación Clientes con Bonificación Especial"
   ClientHeight    =   3915
   ClientLeft      =   45
   ClientTop       =   2430
   ClientWidth     =   6930
   Icon            =   "frmPreFactBonif.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3915
   ScaleWidth      =   6930
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
      Height          =   3885
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   6915
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   2715
         Locked          =   -1  'True
         TabIndex        =   11
         Text            =   "Text5"
         Top             =   2295
         Width           =   3405
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   2715
         Locked          =   -1  'True
         TabIndex        =   10
         Text            =   "Text5"
         Top             =   1920
         Width           =   3405
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   1
         Left            =   1815
         MaxLength       =   6
         TabIndex        =   3
         Top             =   2295
         Width           =   830
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   0
         Left            =   1815
         MaxLength       =   6
         TabIndex        =   2
         Top             =   1920
         Width           =   830
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   3
         Left            =   1815
         MaxLength       =   10
         TabIndex        =   1
         Tag             =   "Código Postal|T|S|||clientes|codposta|||"
         Top             =   1290
         Width           =   1050
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   2
         Left            =   1815
         MaxLength       =   10
         TabIndex        =   0
         Tag             =   "Código Postal|T|S|||clientes|codposta|||"
         Top             =   930
         Width           =   1050
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   5175
         TabIndex        =   5
         Top             =   3300
         Width           =   975
      End
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   3990
         TabIndex        =   4
         Top             =   3300
         Width           =   975
      End
      Begin MSComctlLib.ProgressBar Pb1 
         Height          =   255
         Left            =   570
         TabIndex        =   15
         Top             =   3000
         Visible         =   0   'False
         Width           =   5550
         _ExtentX        =   9790
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Label Label4 
         Caption         =   "Actualizando Albaranes:"
         Height          =   195
         Index           =   0
         Left            =   600
         TabIndex        =   16
         Top             =   2730
         Visible         =   0   'False
         Width           =   1965
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   1
         Left            =   1515
         MouseIcon       =   "frmPreFactBonif.frx":000C
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar Cliente"
         Top             =   2295
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   0
         Left            =   1515
         MouseIcon       =   "frmPreFactBonif.frx":015E
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar Cliente"
         Top             =   1920
         Width           =   240
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
         Left            =   570
         TabIndex        =   14
         Top             =   1680
         Width           =   495
      End
      Begin VB.Label Label4 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   12
         Left            =   930
         TabIndex        =   13
         Top             =   2295
         Width           =   420
      End
      Begin VB.Label Label4 
         Caption         =   "Desde"
         Height          =   195
         Index           =   13
         Left            =   930
         TabIndex        =   12
         Top             =   1920
         Width           =   465
      End
      Begin VB.Label Label4 
         Caption         =   "Fecha Albarán"
         ForeColor       =   &H00972E0B&
         Height          =   255
         Index           =   16
         Left            =   570
         TabIndex        =   9
         Top             =   690
         Width           =   1815
      End
      Begin VB.Label Label4 
         Caption         =   "Desde"
         Height          =   195
         Index           =   15
         Left            =   930
         TabIndex        =   8
         Top             =   930
         Width           =   465
      End
      Begin VB.Label Label4 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   14
         Left            =   930
         TabIndex        =   7
         Top             =   1290
         Width           =   420
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   2
         Left            =   1500
         Picture         =   "frmPreFactBonif.frx":02B0
         ToolTipText     =   "Buscar fecha"
         Top             =   930
         Width           =   240
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   3
         Left            =   1500
         Picture         =   "frmPreFactBonif.frx":033B
         ToolTipText     =   "Buscar fecha"
         Top             =   1290
         Width           =   240
      End
   End
End
Attribute VB_Name = "frmPreFactBonif"
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

Private WithEvents frmC As frmCal 'calendario fechas
Attribute frmC.VB_VarHelpID = -1
Private WithEvents frmcli As frmManClien 'ayuda de socios
Attribute frmcli.VB_VarHelpID = -1

Dim indCodigo As Integer 'indice para txtCodigo
Dim indFrame As Single 'nº de frame en el que estamos
 
Dim PrimeraVez As Boolean
Dim tabla As String
Dim Codigo As String 'Código para FormulaSelection de Crystal Report
Dim TipCod As String
Dim Orden1 As String 'Campo de Ordenacion (por codigo) para Cristal Report
Dim Orden2 As String 'Campo de Ordenacion (por nombre) para Cristal Report


'GENERALES PARA PASARLE A CRYSTAL REPORT
Private cadFormula As String 'Cadena con la FormulaSelection para Crystal Report
Private cadParam As String 'Cadena con los parametros para Crystal Report
Private numParam As Byte 'Numero de parametros que se pasan a Crystal Report
Private cadSelect As String 'Cadena para comprobar si hay datos antes de abrir Informe
Private cadTitulo As String 'Titulo para la ventana frmImprimir
Private cadNombreRPT As String 'Nombre del informe


Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, 0, cerrar
    If cerrar Then Unload Me
End Sub

Private Sub InicializarVbles()
    cadFormula = ""
    cadSelect = ""
    cadParam = ""
    numParam = 0
End Sub



Private Sub cmdAceptar_Click()
Dim cDesde As String, cHasta As String 'cadena codigo Desde/Hasta
Dim nDesde As String, nHasta As String 'cadena Descripcion Desde/Hasta
Dim cadTABLA As String, cOrden As String
Dim i As Byte
Dim sql As String
Dim Sql2 As String


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
    
    'D/H Fecha albaran
    cDesde = Trim(txtCodigo(2).Text)
    cHasta = Trim(txtCodigo(3).Text)
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{" & tabla & ".fecalbar}"
        TipCod = "F"
        If Not PonerDesdeHasta(cDesde, cHasta, "", "", "pDHfechaFac= """) Then Exit Sub
    End If
    
    sql = "select distinct " & vSesion.Codigo & ", scaalb.codartic, sartic.nomartic, sartic.ultpreci "
    Sql2 = "select scaalb.codclave "
    sql = sql & " from ((((scaalb INNER JOIN sartic ON scaalb.codartic = sartic.codartic) "
    Sql2 = Sql2 & " from ((((scaalb INNER JOIN sartic ON scaalb.codartic = sartic.codartic) "
    sql = sql & " INNER JOIN sfamia ON sartic.codfamia = sfamia.codfamia and sfamia.tipfamia = 1) " ' solo articulos de la familia de carburantes
    Sql2 = Sql2 & " INNER JOIN sfamia ON sartic.codfamia = sfamia.codfamia and sfamia.tipfamia = 1) "
    sql = sql & " INNER JOIN ssocio ON scaalb.codsocio = ssocio.codsocio and ssocio.bonifesp = 1) " ' solo clientes que sean de bonificacion especial
    Sql2 = Sql2 & "  INNER JOIN ssocio ON scaalb.codsocio = ssocio.codsocio and ssocio.bonifesp = 1) "
    '[Monica]04/01/2013: Efectivos
    sql = sql & " INNER JOIN sforpa ON scaalb.codforpa = sforpa.codforpa and sforpa.tipforpa <> 0 and sforpa.tipforpa <> 6) " ' solo formas de pago que no sean de efectivo
    Sql2 = Sql2 & " INNER JOIN sforpa ON scaalb.codforpa = sforpa.codforpa and sforpa.tipforpa <> 0 and sforpa.tipforpa <> 6) "
    '[Monica]28/12/2011:tenemos que saber que articulos tienen bonificacion
    sql = sql & " INNER JOIN smargen ON scaalb.codsocio = smargen.codsocio and scaalb.codartic = smargen.codartic " ' solo los articulos que tengan bonificacion
    Sql2 = Sql2 & " INNER JOIN smargen ON scaalb.codsocio = smargen.codsocio and scaalb.codartic = smargen.codartic "
    
    If cadSelect <> "" Then
        sql = sql & " where " & Replace(Replace(cadSelect, "{", ""), "}", "")
        Sql2 = Sql2 & " where " & Replace(Replace(cadSelect, "{", ""), "}", "")
    End If
    
    If TotalRegistrosConsulta(sql) = 0 Then
        MsgBox "No hay albaranes en ese rango. Revise", vbExclamation
        Exit Sub
    End If
    
    '[Monica]15/12/2011: Comprobamos si todos los albaranes no tienen  precio inicial, si hay albaranes damos un aviso
    ' y preguntamos si quiere continuar con el proceso
    If ExistenBonificaciones(sql) Then
        If MsgBox("Existen albaranes que ya están bonificados. " & vbCrLf & vbCrLf & "¿Desea continuar con el proceso?", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
            Exit Sub
        End If
    End If
    
    If ProcesarCambios(sql, Sql2) Then
        cmdCancel_Click
    End If

End Sub

Private Function ProcesarCambios(vSQL As String, vSql2 As String) As Boolean
Dim sql As String
    
    On Error GoTo eProcesarCambios
    
    ProcesarCambios = False
    
    ' cargamos la tabla sobre la que está el formulario para que vean y puedan cambiar precios de ultima compra
    sql = "delete from tmpinformes where codusu = " & vSesion.Codigo
    Conn.Execute sql
    
    sql = "insert into tmpinformes(codusu, codigo1, nombre1, precio2) " & vSQL
    Conn.Execute sql
    
    frmPreciosArt.Show vbModal
    
    If MsgBox("¿Desea continuar con el proceso de modificación de Albaranes?", vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
        If ModificacionAlbaranes(vSql2, "scaalb", Pb1, Label4(0)) Then  'cadSelect) Then
            MsgBox "Proceso realizado correctamente.", vbExclamation
        End If
    End If

    ProcesarCambios = True
    Exit Function
    
eProcesarCambios:
    MuestraError Err.Number, "Procesar Cambios", Err.Description
End Function


Private Function ExistenBonificaciones(vSQL As String) As Boolean
Dim Sql2 As String

    If InStr(1, vSQL, "where") <> 0 Then
        Sql2 = vSQL & " and scaalb.precioinicial <> 0"
    Else
        Sql2 = vSQL & " where scaalb.precioinicial <> 0"
    End If
    ExistenBonificaciones = (TotalRegistrosConsulta(Sql2) <> 0)

End Function



'Private Function ModificacionAlbaranes(cadWhere As String) As Boolean
'Dim sql As String
'Dim sql2 As String
'Dim Sql4 As String
'Dim RS As ADODB.Recordset
'Dim Rs4 As ADODB.Recordset
'Dim Margen As Currency
'Dim EurosLitro As Double
'Dim PrecioNue As Double
'Dim PrecioNue2 As Double
'Dim ImporteNue As Currency
'Dim nRegs As Integer
'
'    On Error GoTo eModificacionAlbaranes
'
'    ModificacionAlbaranes = False
'
'    Conn.BeginTrans
'
'    '[Monica]07/03/2012: cambio del calculo para guardarnos el precio
'    'Sql = "select scaalb.codclave, scaalb.codsocio, scaalb.codartic, scaalb.cantidad, tmpinformes.precio2 "
'    sql = "select distinct scaalb.codsocio, scaalb.codartic, tmpinformes.precio2 "
'    sql = sql & " from scaalb INNER JOIN tmpinformes ON scaalb.codartic = tmpinformes.codigo1 and tmpinformes.codusu = " & vSesion.Codigo
'    sql = sql & " where scaalb.codclave in (" & Replace(Replace(cadWhere, "{", ""), "}", "") & ")"
'
'    nRegs = TotalRegistrosConsulta(sql)
'
'    CargarProgres Pb1, nRegs
'    Pb1.visible = True
'    Label4(0).visible = True
'    DoEvents
'
'    Set RS = New ADODB.Recordset
'    RS.Open sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
'
'    While Not RS.EOF
'        IncrementarProgres Pb1, 1
'        DoEvents
'
'        Margen = DevuelveValor("select margen from smargen where codsocio = " & DBSet(RS!codsocio, "N") & " and codartic = " & DBSet(RS!codartic, "N"))
'        '[Monica]15/12/2011: Euros/litro
'        EurosLitro = DevuelveValor("select euroslitro from smargen where codsocio = " & DBSet(RS!codsocio, "N") & " and codartic = " & DBSet(RS!codartic, "N"))
'
'        If Margen <> 0 Then
'            PrecioNue = CDbl(DBLet(RS!precio2, "N")) * (1 + (Margen / 100))
'        Else
'            PrecioNue = CDbl(DBLet(RS!precio2, "N")) + EurosLitro
'        End If
'
'        PrecioNue2 = Round2(PrecioNue, 3)
'
'        Sql4 = "select scaalb.codclave, scaalb.codsocio, scaalb.codartic, scaalb.cantidad, tmpinformes.precio2"
'        Sql4 = Sql4 & " from scaalb INNER JOIN tmpinformes ON scaalb.codartic = tmpinformes.codigo1 and tmpinformes.codusu = " & vSesion.Codigo
'        Sql4 = Sql4 & " where scaalb.codclave in (" & Replace(Replace(cadWhere, "{", ""), "}", "") & ")"
'        Sql4 = Sql4 & " and scaalb.codsocio = " & DBSet(RS!codsocio, "N")
'        Sql4 = Sql4 & " and scaalb.codartic = " & DBSet(RS!codartic, "N")
'
'        Set Rs4 = New ADODB.Recordset
'        Rs4.Open Sql4, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
'
'        While Not Rs4.EOF
''            Margen = DevuelveValor("select margen from smargen where codsocio = " & DBSet(Rs!codsocio, "N") & " and codartic = " & DBSet(Rs!codartic, "N"))
''            [Monica]15/12/2011: Euros/litro
''            EurosLitro = DevuelveValor("select euroslitro from smargen where codsocio = " & DBSet(Rs!codsocio, "N") & " and codartic = " & DBSet(Rs!codartic, "N"))
''
''            If Margen <> 0 Then
''                PrecioNue = CDbl(DBLet(Rs!precio2, "N")) * (1 + (Margen / 100))
''            Else
''                PrecioNue = CDbl(DBLet(Rs!precio2, "N")) + EurosLitro
''            End If
'            ImporteNue = Round2(PrecioNue * DBLet(Rs4!cantidad, "N"), 2)
'
'            '[Monica]27/02/2012: antes el precio se redondeaba a 3 al calcularlo, ahora lo calculo despues de calcular el importe
''            PrecioNue = Round2(ImporteNue / DBLet(Rs!cantidad, "N"), 3)
'
'            '[Monica]15/12/2011: Precioinicio
'            sql2 = "update scaalb set precioinicial = preciove "
'            sql2 = sql2 & " where codclave = " & DBSet(Rs4!Codclave, "N")
'
'            Conn.Execute sql2
'
'            sql2 = "update scaalb set preciove = " & DBSet(PrecioNue2, "N")
'            sql2 = sql2 & " ,importel = " & DBSet(ImporteNue, "N")
'            sql2 = sql2 & " where codclave = " & DBSet(Rs4!Codclave, "N")
'
'            Conn.Execute sql2
'
'            Rs4.MoveNext
'        Wend
'        Set Rs4 = Nothing
'
'        RS.MoveNext
'    Wend
'
'    Set RS = Nothing
'
'    ModificacionAlbaranes = True
'    Conn.CommitTrans
'    Exit Function
'
'eModificacionAlbaranes:
'    Conn.RollbackTrans
'    MuestraError Err.Number, "Modificacion Albaranes", Err.Description
'End Function
'
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

    Me.imgBuscar(0).Picture = frmPpal.imgListImages16.ListImages(1).Picture
    Me.imgBuscar(1).Picture = frmPpal.imgListImages16.ListImages(1).Picture

 
    '###Descomentar
'    CommitConexion
         
    tabla = "scaalb"
         
         
    FrameCobrosVisible True, h, w
    indFrame = 5
            
    'Esto se consigue poneinedo el cancel en el opcion k corresponda
    Me.cmdCancel.Cancel = True
'    Me.Width = w + 70
'    Me.Height = h + 350
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

Private Sub imgBuscar_Click(Index As Integer)
   Select Case Index
        Case 0, 1 'CLIENTE
            AbrirFrmClientes (Index)
        
   End Select
   PonerFoco txtCodigo(indCodigo)
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

Private Sub AbrirFrmClientes(indice As Integer)
    indCodigo = indice
    Set frmcli = New frmManClien
    frmcli.DatosADevolverBusqueda = "0|1|"
    frmcli.DeConsulta = True
    frmcli.CodigoActual = txtCodigo(indCodigo)
    frmcli.Show vbModal
    Set frmcli = Nothing
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
    KEYpress KeyAscii
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
            
    End Select
End Sub

Private Sub FrameCobrosVisible(visible As Boolean, ByRef h As Integer, ByRef w As Integer)
    Me.FrameCobros.visible = visible
    If visible = True Then
        Me.FrameCobros.Top = -90
        Me.FrameCobros.Left = 0
        Me.FrameCobros.Height = 4305
        Me.FrameCobros.Width = 7020
        w = Me.FrameCobros.Width
        h = Me.FrameCobros.Height
    End If
End Sub

Private Sub CargarTablaTemporal(DesFec As String, HasFec As String, SoloDesc As Byte)
    Dim sql As String
    Dim SQL1 As String
    Dim Sql2 As String
    Dim Sql3 As String
    Dim RS As ADODB.Recordset
    Dim Rs1 As ADODB.Recordset
    Dim Rs2 As ADODB.Recordset

    On Error GoTo eCargarTablaTemporal

    ' primero borramos los registros del usuario
    sql = "delete from tmpinformes where codusu = " & vSesion.Codusu
    Conn.Execute sql
    

    ' cargamos la tabla temporal para el listado agrupando por fecha y turno
    ' unicamente cargamos el importe de mangueras, el resto lo inicializamos a 0
    sql = "select fechatur, codturno, sum(importel) from sturno where "
    If DesFec <> "" Then
        sql = sql & " fechatur >= '" & Format(DesFec, FormatoFecha) & "'"
    End If
    If HasFec <> "" Then
        sql = sql & " and fechatur <= '" & Format(HasFec, FormatoFecha) & "'"
    End If
    sql = sql & " and tipocred = 0 "
    sql = sql & " group by fechatur, codturno "
    sql = sql & " order by fechatur, codturno "
    
    Set RS = New ADODB.Recordset ' Crear objeto
    RS.Open sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText ' abrir cursor
      
    If Not RS.EOF Then RS.MoveFirst
    
    While Not RS.EOF
        SQL1 = "insert into tmpinformes (codusu, fecha1, campo1, importe1, importe2, importe3, importe4) "
        SQL1 = SQL1 & "values (" & vSesion.Codigo & ",'" & Format(RS.Fields(0).Value, FormatoFecha) & "',"
        If Not IsNull(RS.Fields(2).Value) Then
            SQL1 = SQL1 & RS.Fields(1).Value & "," & TransformaComasPuntos(ImporteSinFormato(RS.Fields(2).Value)) & ",0,0,0)"
        Else
            SQL1 = SQL1 & RS.Fields(1).Value & ",0,0,0,0)"
        End If
        
        Conn.Execute SQL1
    
        RS.MoveNext
    Wend
    
    RS.Close
    sql = "select fecha1, campo1 from tmpinformes where codusu = " & vSesion.Codigo
    sql = sql & " order by 1, 2"
    
    RS.Open sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText ' abrir cursor
    If Not RS.EOF Then RS.MoveFirst
    While Not RS.EOF
        SQL1 = "select sum(importel) from scaalb where fecalbar = '" & Format(RS.Fields(0).Value, FormatoFecha)
        SQL1 = SQL1 & "' and codturno = " & RS.Fields(1).Value & " and codartic >=1 and codartic <= 9 "
    
        Set Rs1 = New ADODB.Recordset
        Rs1.Open SQL1, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText ' abrir cursor
        If Not Rs1.EOF Then Rs1.MoveFirst
        
        Sql2 = "select sum(importel) from scaalb where fecalbar = '" & Format(RS.Fields(0).Value, FormatoFecha)
        Sql2 = Sql2 & "' and codturno = " & RS.Fields(1).Value & " and codartic >=1 and codartic <= 9 "
        Sql2 = Sql2 & " and numalbar = 'MANUAL'"
        
        Set Rs2 = New ADODB.Recordset
        Rs2.Open Sql2, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText ' abrir cursor
        If Not Rs2.EOF Then Rs2.MoveFirst
        
        
        Sql3 = "update tmpinformes set importe2 = "
        If Not IsNull(Rs1.Fields(0).Value) Then
            Sql3 = Sql3 & TransformaComasPuntos(ImporteSinFormato(Rs1.Fields(0).Value)) & ", "
        Else
            Sql3 = Sql3 & "0,"
        End If
        
        If Not IsNull(Rs2.Fields(0).Value) Then
            Sql3 = Sql3 & "importe4 = " & TransformaComasPuntos(ImporteSinFormato(Rs2.Fields(0).Value))
        Else
            Sql3 = Sql3 & "importe4 = 0 "
        End If
        
        Sql3 = Sql3 & " where fecha1 = '" & Format(RS.Fields(0).Value, FormatoFecha) & "' and "
        Sql3 = Sql3 & " campo1 = " & RS.Fields(1).Value
        Sql3 = Sql3 & " and codusu = " & vSesion.Codigo
        
        
        Conn.Execute Sql3
        
        Set Rs1 = Nothing
        Set Rs2 = Nothing
        
        Debug.Print RS.Fields(0).Value & "-" & RS.Fields(1).Value
        
        RS.MoveNext
    Wend

    ' una vez cargada la tabla temporal acualizamos el importe3 = diferencia entre importe1 e importe2
    sql = "update tmpinformes set importe3 = importe1 - importe2 where codusu = " & vSesion.Codigo
    Conn.Execute sql

    If SoloDesc = 1 Then
        sql = "delete from tmpinformes where codusu = " & vSesion.Codigo
        sql = sql & " and importe3 > -1 and importe3 < 1 "
        
        Conn.Execute sql
    End If

eCargarTablaTemporal:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Error en la carga de la tabla temporal"
    End If
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

