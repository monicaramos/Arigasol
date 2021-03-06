VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmEstVtasdia 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Resumen Ventas diarias"
   ClientHeight    =   3585
   ClientLeft      =   45
   ClientTop       =   2430
   ClientWidth     =   4965
   Icon            =   "frmEstVtasdia.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3585
   ScaleWidth      =   4965
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
      Height          =   3375
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   4755
      Begin VB.Frame Frame4 
         ForeColor       =   &H00972E0B&
         Height          =   795
         Left            =   600
         TabIndex        =   10
         Top             =   1710
         Width           =   2175
         Begin VB.OptionButton Option1 
            Caption         =   "Albaranes"
            Height          =   255
            Index           =   0
            Left            =   210
            TabIndex        =   2
            Top             =   180
            Value           =   -1  'True
            Width           =   1815
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Facturas"
            Height          =   255
            Index           =   1
            Left            =   210
            TabIndex        =   3
            Top             =   420
            Width           =   1815
         End
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   3
         Left            =   1845
         MaxLength       =   10
         TabIndex        =   1
         Tag             =   "C�digo Postal|T|S|||clientes|codposta|||"
         Top             =   1200
         Width           =   1050
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   2
         Left            =   1845
         MaxLength       =   10
         TabIndex        =   0
         Tag             =   "C�digo Postal|T|S|||clientes|codposta|||"
         Top             =   840
         Width           =   1050
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   3375
         TabIndex        =   5
         Top             =   2700
         Width           =   975
      End
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   2190
         TabIndex        =   4
         Top             =   2700
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "Fecha"
         ForeColor       =   &H00972E0B&
         Height          =   255
         Index           =   16
         Left            =   600
         TabIndex        =   9
         Top             =   600
         Width           =   1815
      End
      Begin VB.Label Label4 
         Caption         =   "Desde"
         Height          =   195
         Index           =   15
         Left            =   960
         TabIndex        =   8
         Top             =   840
         Width           =   465
      End
      Begin VB.Label Label4 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   14
         Left            =   960
         TabIndex        =   7
         Top             =   1200
         Width           =   420
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   2
         Left            =   1530
         Picture         =   "frmEstVtasdia.frx":000C
         ToolTipText     =   "Buscar fecha"
         Top             =   840
         Width           =   240
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   3
         Left            =   1530
         Picture         =   "frmEstVtasdia.frx":0097
         ToolTipText     =   "Buscar fecha"
         Top             =   1200
         Width           =   240
      End
   End
End
Attribute VB_Name = "frmEstVtasdia"
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
Dim cadTABLA As String, cOrden As String
Dim i As Byte
InicializarVbles
    
    '========= PARAMETROS  =============================
    'A�adir el parametro de Empresa
    cadParam = cadParam & "|pEmpresa=""" & vEmpresa.nomEmpre & """|"
    numParam = numParam + 1
    
    'D/H Fecha factura
    cDesde = Trim(txtCodigo(2).Text)
    cHasta = Trim(txtCodigo(3).Text)
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        If Option1(1).Value Then Codigo = "{" & tabla & ".fecfactu}"
        If Option1(0).Value Then Codigo = "{slhfac.fecalbar}"
        TipCod = "F"
        If Not PonerDesdeHasta(cDesde, cHasta, "", "", "pDHfechaFac= """) Then Exit Sub
    End If
    
    'Comprobar si hay registros a Mostrar antes de abrir el Informe
    If Option1(1).Value Then cadTABLA = "schfac"
    If Option1(0).Value Then cadTABLA = "slhfac"
'    cadTABLA = tabla
    
    If HayRegParaInforme(cadTABLA, cadSelect) Then
        If Option1(1).Value Then
            cadTitulo = "Resumen Ventas diarias-Facturas"
            cadNombreRPT = "rEstVtasdia.rpt"
        End If
        If Option1(0).Value Then
            cadTitulo = "Resumen Ventas diarias-Albaranes"
            cadFormula = ""
            AnyadirAFormula cadFormula, "{tmpinformes.codusu} = " & Format(vSesion.Codigo, "0000")
            CargarTablaIntermedia
            cadNombreRPT = "rEstVtasdiaAlb.rpt"
        End If
       LlamarImprimir
       'AbrirVisReport
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
    Limpiar Me

    '###Descomentar
'    CommitConexion
         
    FrameCobrosVisible True, h, w
    indFrame = 5
    tabla = "schfac"
            
    'Esto se consigue poneinedo el cancel en el opcion k corresponda
    Me.cmdCancel.Cancel = True
'    Me.Width = w + 70
'    Me.Height = h + 350
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
'15/02/2007
'    KEYpress KeyAscii
'ahora
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
            
    End Select
End Sub
Private Sub FrameCobrosVisible(visible As Boolean, ByRef h As Integer, ByRef w As Integer)
'Frame para los cobros a clientes por fecha vencimiento
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
        .Opcion = 0
        .Show vbModal
    End With
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

Private Function CargarTablaIntermedia() As Boolean
Dim Sql As String
Dim Rs As ADODB.Recordset
Dim AntFecha As Date
Dim AntIva As Integer
Dim ActFecha As Date
Dim ActIva As Integer
Dim vBase As Currency
Dim vCuota As Currency
Dim vImporte As Currency
Dim PorcIva As String
Dim cuota As Currency
Dim base As Currency
Dim sql2 As String
Dim HayReg As Byte

    On Error GoTo eCargarTablaIntermedia

    CargarTablaIntermedia = False

    Sql = "delete from tmpinformes where codusu = " & vSesion.Codigo
    Conn.Execute Sql


    Sql = "select slhfac.fecalbar, sartic.codigiva, sum(slhfac.implinea) from slhfac, sartic where 1 = 1 "
    If txtCodigo(2).Text <> "" Then Sql = Sql & " and slhfac.fecalbar >= " & DBSet(txtCodigo(2).Text, "F")
    If txtCodigo(3).Text <> "" Then Sql = Sql & " and slhfac.fecalbar <= " & DBSet(txtCodigo(3).Text, "F")

    Sql = Sql & " and slhfac.codartic = sartic.codartic "
    Sql = Sql & " group by 1,2"
    Sql = Sql & " order by 1,2"

    Set Rs = New ADODB.Recordset
    Rs.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText

    ActFecha = DBLet(Rs.Fields(0).Value, "F")
    AntFecha = ActFecha
    If Not Rs.EOF Then Rs.MoveFirst
    
    vBase = 0
    vCuota = 0
    vImporte = 0
    
    While Not Rs.EOF
        HayReg = 1
        ActFecha = DBLet(Rs.Fields(0).Value, "F")
        
        If ActFecha <> AntFecha Then
              ' hacemos es insert en la temporal
              sql2 = "insert into tmpinformes (codusu, fecha1, importe1, importe2, importe3) values ("
              sql2 = sql2 & vSesion.Codigo & "," & DBSet(AntFecha, "F") & "," & DBSet(vBase, "N") & ","
              sql2 = sql2 & DBSet(vCuota, "N") & "," & DBSet(vImporte, "N") & ")"

              Conn.Execute sql2

              vBase = 0
              vCuota = 0
              vImporte = 0

              AntFecha = ActFecha
        End If

        PorcIva = ""
        PorcIva = DevuelveDesdeBDNew(cConta, "tiposiva", "porceiva", "codigiva", Rs.Fields(1).Value, "N")
        base = Round2(DBLet(Rs.Fields(2).Value, "N") / (1 + CCur(PorcIva) / 100), 2)
        cuota = DBLet(Rs.Fields(2).Value, "N") - base

        vBase = vBase + base
        vCuota = vCuota + cuota
        vImporte = vImporte + DBLet(Rs.Fields(2).Value, "N")

        Rs.MoveNext

    Wend
    If HayReg = 1 Then
        sql2 = "insert into tmpinformes (codusu, fecha1, importe1, importe2, importe3) values ("
        sql2 = sql2 & vSesion.Codigo & "," & DBSet(ActFecha, "F") & "," & DBSet(vBase, "N") & ","
        sql2 = sql2 & DBSet(vCuota, "N") & "," & DBSet(vImporte, "N") & ")"

        Conn.Execute sql2
    End If
    Set Rs = Nothing
    CargarTablaIntermedia = True

eCargarTablaIntermedia:
    If Err.Number <> 0 Then
        MsgBox "Error cargando la tabla intermedia " & Err.Description, vbExclamation
    End If

End Function

