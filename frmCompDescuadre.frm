VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmCompDescuadre 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Comprobar Descuadres Turno"
   ClientHeight    =   3645
   ClientLeft      =   45
   ClientTop       =   2430
   ClientWidth     =   7185
   Icon            =   "frmCompDescuadre.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3645
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
      Height          =   3405
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   6915
      Begin VB.CheckBox Check1 
         Caption         =   "Imprimir sólo descuadres"
         Height          =   255
         Left            =   600
         TabIndex        =   2
         Top             =   2010
         Width           =   2445
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
         Left            =   4905
         TabIndex        =   4
         Top             =   2550
         Width           =   975
      End
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   3660
         TabIndex        =   3
         Top             =   2550
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "Fecha factura"
         ForeColor       =   &H00972E0B&
         Height          =   255
         Index           =   16
         Left            =   570
         TabIndex        =   8
         Top             =   690
         Width           =   1815
      End
      Begin VB.Label Label4 
         Caption         =   "Desde"
         Height          =   195
         Index           =   15
         Left            =   930
         TabIndex        =   7
         Top             =   930
         Width           =   465
      End
      Begin VB.Label Label4 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   14
         Left            =   930
         TabIndex        =   6
         Top             =   1290
         Width           =   420
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   2
         Left            =   1500
         Picture         =   "frmCompDescuadre.frx":000C
         ToolTipText     =   "Buscar fecha"
         Top             =   930
         Width           =   240
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   3
         Left            =   1500
         Picture         =   "frmCompDescuadre.frx":0097
         ToolTipText     =   "Buscar fecha"
         Top             =   1290
         Width           =   240
      End
   End
End
Attribute VB_Name = "frmCompDescuadre"
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

    CargarTablaTemporal Trim(txtCodigo(2).Text), Trim(txtCodigo(3).Text), Check1.Value


    InicializarVbles
    
    '========= PARAMETROS  =============================
    'Añadir el parametro de Empresa
    cadParam = cadParam & "|pEmpresa=""" & vEmpresa.nomEmpre & """|"
    numParam = numParam + 1
    
    cadParam = cadParam & ""
    
    'AL usar una tabla temporal tenemos que indicar el usuario que toca
    Codigo = "{tmpinformes.codusu}"
    cDesde = Codigo & " = " & vSesion.Codigo
    cDesde = "(" & cDesde & ")"
    
    AnyadirAFormula cadFormula, cDesde

'    cadFormula = "{tmpinformes.codusu} =" & vSesion.Empleado
    
    'D/H Fecha factura
    cDesde = Trim(txtCodigo(2).Text)
    cHasta = Trim(txtCodigo(3).Text)
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{tmpinformes.fecha1}"
        TipCod = "F"
        If Not PonerDesdeHasta(cDesde, cHasta, "", "", "pDHfechaFac= """) Then Exit Sub
    End If
    
'    cadParam = cadParam & "pSoloDesc=" & Format(Check1.Value, "0") & "|"
'    numParam = numParam + 1
    cadTABLA = tabla

    If HayRegParaInforme(cadTABLA, cadSelect) Then
       cadTitulo = "Comprobación descuadres turno"
       cadNombreRPT = "rCompDescu2.rpt"
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
        PonerFoco txtCodigo(0)
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
Dim h As Integer, w As Integer
Dim List As Collection

    PrimeraVez = True
    limpiar Me

 
    '###Descomentar
'    CommitConexion
         
    FrameCobrosVisible True, h, w
    indFrame = 5
    tabla = "tmpinformes"
            
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
'14/02/2007 antes era esto
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

Private Sub CargarTablaTemporal(DesFec As String, HasFec As String, SoloDesc As Byte)
    Dim sql As String
    Dim SQL1 As String
    Dim Sql2 As String
    Dim sql3 As String
    Dim RS As ADODB.Recordset
    Dim Rs1 As ADODB.Recordset
    Dim Rs2 As ADODB.Recordset
    Dim Importe As Currency

    On Error GoTo eCargarTablaTemporal

    ' primero borramos los registros del usuario
    sql = "delete from tmpinformes where codusu = " & DBSet(vSesion.Codigo, "N")
    Conn.Execute sql

    '[Monica]11/01/2013: Comprobamos que no sea Ribarroja
    If vParamAplic.Cooperativa = 5 Then Exit Sub

    ' cargamos la tabla temporal para el listado agrupando por fecha y turno
    ' unicamente cargamos el importe de mangueras, el resto lo inicializamos a 0
    sql = "select fechatur, codturno, sum(importel) from sturno where "
    If DesFec <> "" Then
        sql = sql & " fechatur >= " & DBSet(DesFec, "F")
    End If
    If HasFec <> "" Then
        sql = sql & " and fechatur <= " & DBSet(HasFec, "F")
    End If
    sql = sql & " and tipocred = 0 "
    sql = sql & " group by fechatur, codturno "
    sql = sql & " order by fechatur, codturno "
    
    Set RS = New ADODB.Recordset ' Crear objeto
    RS.Open sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText ' abrir cursor
      
    If Not RS.EOF Then RS.MoveFirst
    
    While Not RS.EOF
        SQL1 = "insert into tmpinformes (codusu, fecha1, campo1, importe1, importe2, importe3, importe4) "
        SQL1 = SQL1 & "values (" & DBSet(vSesion.Codigo, "N") & "," & DBSet(RS.Fields(0).Value, "F") & ","
        SQL1 = SQL1 & DBLet(RS.Fields(1).Value, "N") & ","
        
        Importe = DBLet(RS.Fields(2).Value, "N")
        SQL1 = SQL1 & TransformaComasPuntos(ImporteSinFormato(CStr(Importe))) & ",0,0,0)"
        
        Conn.Execute SQL1
    
        RS.MoveNext
    Wend
    
    RS.Close
    sql = "select fecha1, campo1 from tmpinformes where codusu = " & DBSet(vSesion.Codigo, "N")
    sql = sql & " order by 1, 2"
    
    RS.Open sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText ' abrir cursor
    If Not RS.EOF Then RS.MoveFirst
    While Not RS.EOF
        SQL1 = "select sum(importel) from scaalb where fecalbar = '" & DBLet(RS.Fields(0).Value, "F")
        SQL1 = SQL1 & "' and codturno = " & DBLet(RS.Fields(1).Value, "N") & " and codartic >=1 and codartic <= 9 "
    
        Set Rs1 = New ADODB.Recordset
        Rs1.Open SQL1, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText ' abrir cursor
        If Not Rs1.EOF Then Rs1.MoveFirst
        
        Sql2 = "select sum(importel) from scaalb where fecalbar = '" & DBLet(RS.Fields(0).Value, "F")
        Sql2 = Sql2 & "' and codturno = " & DBLet(RS.Fields(1).Value, "N") & " and codartic >=1 and codartic <= 9 "
        Sql2 = Sql2 & " and numalbar = 'MANUAL'"
        
        Set Rs2 = New ADODB.Recordset
        Rs2.Open Sql2, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText ' abrir cursor
        If Not Rs2.EOF Then Rs2.MoveFirst
        
        
        sql3 = "update tmpinformes set importe2 = "
        Importe = DBLet(Rs1.Fields(0).Value, "N")
        sql3 = sql3 & TransformaComasPuntos(ImporteSinFormato(CStr(Importe))) & ", "
        
        Importe = DBLet(Rs2.Fields(0).Value, "N")
        sql3 = sql3 & "importe4 = " & TransformaComasPuntos(ImporteSinFormato(CStr(Importe)))
        
        sql3 = sql3 & " where fecha1 = '" & DBLet(RS.Fields(0).Value, "F") & "' and "
        sql3 = sql3 & " campo1 = " & DBLet(RS.Fields(1).Value, "N")
        sql3 = sql3 & " and codusu = " & vSesion.Codigo
        
        Conn.Execute sql3
        
        Set Rs1 = Nothing
        Set Rs2 = Nothing
        
'        Debug.Print RS.Fields(0).Value & "-" & RS.Fields(1).Value
        
        RS.MoveNext
    Wend

    ' una vez cargada la tabla temporal acualizamos el importe3 = diferencia entre importe1 e importe2
    sql = "update tmpinformes set importe3 = importe1 - importe2 where codusu = " & DBSet(vSesion.Codigo, "N")
    Conn.Execute sql

    If SoloDesc = 1 Then
        sql = "delete from tmpinformes where codusu = " & DBSet(vSesion.Codigo, "N")
        sql = sql & " and importe3 > -1 and importe3 < 1 "
        
        Conn.Execute sql
    End If

eCargarTablaTemporal:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Error en la carga de la tabla temporal"
    End If
End Sub
