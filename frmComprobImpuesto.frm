VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmComprobImpuesto 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Comprobación de Impuesto en Facturas"
   ClientHeight    =   3000
   ClientLeft      =   45
   ClientTop       =   2430
   ClientWidth     =   6645
   Icon            =   "frmComprobImpuesto.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3000
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
      Height          =   2745
      Left            =   90
      TabIndex        =   4
      Top             =   120
      Width           =   6375
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   3
         Left            =   3735
         MaxLength       =   10
         TabIndex        =   1
         Tag             =   "Código Postal|T|S|||clientes|codposta|||"
         Top             =   840
         Width           =   1050
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   2
         Left            =   1605
         MaxLength       =   10
         TabIndex        =   0
         Tag             =   "Código Postal|T|S|||clientes|codposta|||"
         Top             =   810
         Width           =   1050
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   4755
         TabIndex        =   3
         Top             =   2070
         Width           =   975
      End
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   3570
         TabIndex        =   2
         Top             =   2070
         Width           =   975
      End
      Begin MSComctlLib.ProgressBar Pb1 
         Height          =   255
         Left            =   330
         TabIndex        =   8
         Top             =   1500
         Width           =   5370
         _ExtentX        =   9472
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Label lblProgres 
         Height          =   285
         Index           =   1
         Left            =   330
         TabIndex        =   10
         Top             =   4620
         Width           =   5295
      End
      Begin VB.Label lblProgres 
         Height          =   285
         Index           =   0
         Left            =   360
         TabIndex        =   9
         Top             =   4260
         Width           =   5265
      End
      Begin VB.Label Label4 
         Caption         =   "Fecha Factura"
         ForeColor       =   &H00972E0B&
         Height          =   255
         Index           =   16
         Left            =   330
         TabIndex        =   7
         Top             =   570
         Width           =   1815
      End
      Begin VB.Label Label4 
         Caption         =   "Desde"
         Height          =   195
         Index           =   15
         Left            =   720
         TabIndex        =   6
         Top             =   810
         Width           =   465
      End
      Begin VB.Label Label4 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   14
         Left            =   2880
         TabIndex        =   5
         Top             =   840
         Width           =   420
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   2
         Left            =   1290
         Picture         =   "frmComprobImpuesto.frx":000C
         ToolTipText     =   "Buscar fecha"
         Top             =   810
         Width           =   240
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   3
         Left            =   3420
         Picture         =   "frmComprobImpuesto.frx":0097
         ToolTipText     =   "Buscar fecha"
         Top             =   840
         Width           =   240
      End
   End
End
Attribute VB_Name = "frmComprobImpuesto"
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
Dim sql As String
Dim tipo As Byte
Dim nRegs As Long
Dim NumError As Long

InicializarVbles
    
    '========= PARAMETROS  =============================
    'Añadir el parametro de Empresa
    cadParam = cadParam & "|pEmpresa=""" & vEmpresa.nomEmpre & """|"
    numParam = numParam + 1

    If Not DatosOk Then Exit Sub
    
    'D/H Fecha factura
    cDesde = Trim(txtCodigo(2).Text)
    cHasta = Trim(txtCodigo(3).Text)
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{" & tabla & ".fecfactu}"
        TipCod = "F"
        If Not PonerDesdeHasta(cDesde, cHasta, "", "", "pDHfechaFac= """) Then Exit Sub
    End If
    
    If Not HayRegParaInforme(tabla, cadSelect) Then Exit Sub
    
    ComprobarImpuesto
     
     
eError:
    If Err.Number <> 0 Or NumError <> 0 Then
        MuestraError Err.Number, "No se ha realizado el proceso de comprobación. Llame a soporte."
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
        PonerFoco txtCodigo(2)
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
    tabla = "schfac"
    
    Pb1.visible = False
    
    'Esto se consigue poneinedo el cancel en el opcion k corresponda
    Me.cmdCancel.Cancel = True
    Me.Width = w + 70
    Me.Height = h + 350
    
'14/02/2007 lo dejo donde estaba
'
'     '07022007   lo he quitado de contbilizar facturas y lo he puesto aqui para acotar registros
'     'comprobar que se han rellenado los dos campos de fecha
'     'sino rellenar con fechaini o fechafin del ejercicio
'     'que guardamos en vbles Orden1,Orden2
'
'     Orden1 = ""
'     Orden1 = DevuelveDesdeBDNew(cConta, "parametros", "fechaini", "", "", "", "", "", "", "", "", "", "")
'
'     Orden2 = ""
'     Orden2 = DevuelveDesdeBDNew(cConta, "parametros", "fechafin", "", "", "", "", "", "", "", "", "", "")
     
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
        
    End Select
End Sub

Private Sub FrameCobrosVisible(visible As Boolean, ByRef h As Integer, ByRef w As Integer)
    Me.FrameCobros.visible = visible
    If visible = True Then
        Me.FrameCobros.Top = -90
        Me.FrameCobros.Left = 0
        Me.FrameCobros.Height = 2745
        Me.FrameCobros.Width = 6555
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

' copiado del ariges
Private Sub ComprobarImpuesto()
'Contabiliza Facturas de Clientes o de Proveedores
Dim sql As String
Dim b As Boolean
Dim tmpErrores As Boolean 'Indica si se creo correctamente la tabla de errores
Dim CCoste As String


    sql = "COMIMP"
    'Bloquear para que nadie mas pueda contabilizar

    If Not BloqueoManual(sql, "1") Then
        MsgBox "No se pueden Comprobar Impuesto de Facturas. Hay otro usuario realizando el proceso.", vbExclamation
        Screen.MousePointer = vbDefault
        Exit Sub
    End If

'14/02/2007 lo he descomentado
     'comprobar que se han rellenado los dos campos de fecha
     'sino rellenar con fechaini o fechafin del ejercicio
     'que guardamos en vbles Orden1,Orden2

     Orden1 = ""
     Orden1 = DevuelveDesdeBDNew(cConta, "parametros", "fechaini", "", "", "", "", "", "", "", "", "", "")

     Orden2 = ""
     Orden2 = DevuelveDesdeBDNew(cConta, "parametros", "fechafin", "", "", "", "", "", "", "", "", "", "")

     If txtCodigo(2).Text = "" Then
        txtCodigo(2).Text = Orden1 'fechaini del ejercicio de la conta
     End If

     If txtCodigo(3).Text = "" Then
        txtCodigo(3).Text = Orden2 'fecha fin del ejercicio de la conta
     End If
'14/02/2007 hasta aqui lo he descomentado

    
    'Visualizar la barra de Progreso
    Me.Pb1.visible = True
    
    '===========================================================================
    'COMPROBACION DE IMPUESTO
    '===========================================================================
    Me.lblProgres(0).Caption = "Comprobación de Impuesto en Facturas: "
       
    'vaciamos previamente la temporal
    sql = "delete from tmpinformes where codusu = " & vSesion.Codigo
    Conn.Execute sql
       
    b = ComprobImpuesto()
    
    If b Then
        sql = "select count(*) from tmpinformes where codusu = " & vSesion.Codigo
        If TotalRegistros(sql) <> 0 Then
            cadFormula = ""
            b = AnyadirAFormula(cadFormula, "{tmpinformes.codusu} = " & vSesion.Codigo)
            
            cadTitulo = "Facturas con Impuesto Erroneo"
            cadNombreRPT = "rFacErr.rpt"
            LlamarImprimir
            If MsgBox("Desea corregir el impuesto de las facturas", vbQuestion + vbYesNo + vbDefaultButton1) = vbYes Then
                b = CorregirImpuesto
                If b Then
                    MsgBox "Proceso realizado correctamente.", vbExclamation
                End If
            End If
        Else
            MsgBox "No hay facturas con impuesto erróneo.", vbExclamation
        End If
    End If
    
    'Desbloqueamos ya no estamos comprobando impuestos
    DesBloqueoManual ("COMIMP") 'COMprobar IMPuesto
    
End Sub


Private Function ComprobImpuesto() As Boolean
Dim sql As String
Dim Sql2 As String
Dim sql3 As String
Dim SQL1 As String

Dim Rs As ADODB.Recordset
Dim b As Boolean
Dim i As Integer
Dim numfactu As Long
Dim codigo1 As String
Dim Impuesto As Currency
Dim Hayreg As Boolean
Dim TotalImp As Currency

Dim AntLetraser As String
Dim AntNumfactu As Long
Dim AntFecfactu As Date

Dim ActLetraser As String
Dim ActNumfactu As Long
Dim ActFecfactu As Date

    On Error GoTo EComprobImpuesto

    ComprobImpuesto = False
    
    'Total de Facturas a Insertar en la contabilidad
    sql = "SELECT count(*) "
    sql = sql & " FROM slhfac where 1 = 1 "
    If txtCodigo(2).Text <> "" Then sql = sql & " and fecfactu >= " & DBSet(txtCodigo(2).Text, "F")
    If txtCodigo(3).Text <> "" Then sql = sql & " and fecfactu <= " & DBSet(txtCodigo(3).Text, "F")
    
    Set Rs = New ADODB.Recordset
    Rs.Open sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not Rs.EOF Then
        numfactu = Rs.Fields(0)
    Else
        numfactu = 0
    End If
    Rs.Close
    Set Rs = Nothing

    If numfactu > 0 Then
        CargarProgres Me.Pb1, numfactu
        
        sql = "SELECT * "
        sql = sql & " FROM slhfac where 1 = 1 "
        If txtCodigo(2).Text <> "" Then sql = sql & " and fecfactu >= " & DBSet(txtCodigo(2).Text, "F")
        If txtCodigo(3).Text <> "" Then sql = sql & " and fecfactu <= " & DBSet(txtCodigo(3).Text, "F")
        sql = sql & " order by letraser, numfactu, fecfactu "
            
        Set Rs = New ADODB.Recordset
        Rs.Open sql, Conn, adOpenStatic, adLockPessimistic, adCmdText
        i = 1

        b = True
        
        If Not Rs.EOF Then
            ActLetraser = DBLet(Rs!letraser, "T")
            ActNumfactu = DBLet(Rs!numfactu, "N")
            ActFecfactu = DBLet(Rs!fecfactu, "F")
            
            AntLetraser = ActLetraser
            AntNumfactu = ActNumfactu
            AntFecfactu = ActFecfactu
        End If
        
        'contabilizar cada una de las facturas seleccionadas
        Hayreg = False
        While Not Rs.EOF
            Hayreg = True
            ActLetraser = DBLet(Rs!letraser, "T")
            ActNumfactu = DBLet(Rs!numfactu, "N")
            ActFecfactu = DBLet(Rs!fecfactu, "F")
            
            If ActLetraser <> AntLetraser Or ActNumfactu <> AntNumfactu Or ActFecfactu <> AntFecfactu Then
                ' miramos el resultado en cabecera
                Sql2 = ""
                Sql2 = DevuelveDesdeBDNew(cPTours, "schfac", "impuesto", "letraser", AntLetraser, "T", , "numfactu", CStr(AntNumfactu), "N", "fecfactu", CStr(AntFecfactu), "F")
                
                If TotalImp <> CCur(Sql2) Then
                    sql3 = "insert into tmpinformes (codusu, nombre1, importe1, fecha1, importe2, importe3) values ("
                    sql3 = sql3 & vSesion.Codigo & "," & DBSet(AntLetraser, "T") & "," & DBSet(AntNumfactu, "N") & "," & DBSet(AntFecfactu, "F") & "," & DBSet(Sql2, "N") & "," & DBSet(TotalImp, "N") & ")"
                    
                    Conn.Execute sql3
                End If
                
                TotalImp = 0
                AntNumfactu = ActNumfactu
                AntLetraser = ActLetraser
                AntFecfactu = ActFecfactu
            End If
            
            ' tenemos que calcular el impuesto multiplicando cantidad de linea por impuesto por articulo
            SQL1 = ""
            SQL1 = DevuelveDesdeBD("impuesto", "sartic", "codartic", DBLet(Rs!codArtic), "N")
            If SQL1 = "" Then
                Impuesto = 0
            Else
                Impuesto = CCur(SQL1) ' Comprueba si es nulo y lo pone a 0 o ""
            End If
            
            If EsArticuloCombustible(Rs!codArtic) Then
                TotalImp = TotalImp + Round2((Rs!cantidad * Impuesto), 2)
            End If
            
            IncrementarProgres Me.Pb1, 1
            Me.Refresh
            
            Rs.MoveNext
            
        Wend
        If Hayreg Then
                Sql2 = ""
                Sql2 = DevuelveDesdeBDNew(cPTours, "schfac", "impuesto", "letraser", ActLetraser, "T", , "numfactu", CStr(ActNumfactu), "N", "fecfactu", CStr(ActFecfactu), "F")
                
                If TotalImp <> CCur(Sql2) Then
                    sql3 = "insert into tmpinformes (codusu, nombre1, importe1, fecha1, importe2, importe3) values ("
                    sql3 = sql3 & vSesion.Codigo & "," & DBSet(ActLetraser, "T") & "," & DBSet(ActNumfactu, "N") & "," & DBSet(ActFecfactu, "F") & "," & DBSet(Sql2, "N") & "," & DBSet(TotalImp, "N") & ")"
                    
                    Conn.Execute sql3
                End If
                
                TotalImp = 0
        End If
        
        Rs.Close
        Set Rs = Nothing
    End If
    
EComprobImpuesto:
    If Err.Number = 0 Then
        ComprobImpuesto = True
    Else
        ComprobImpuesto = False
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

Private Function CorregirImpuesto() As Boolean
Dim sql As String
Dim Sql2 As String
Dim Rs As ADODB.Recordset
    On Error GoTo eCorregirImpuesto
    CorregirImpuesto = False
    
    sql = "select * from tmpinformes where codusu = " & vSesion.Codigo
    Set Rs = New ADODB.Recordset
    Rs.Open sql, Conn, adOpenStatic, adLockPessimistic, adCmdText
    
    If Not Rs.EOF Then Rs.MoveFirst
    
    While Not Rs.EOF
        Sql2 = "update schfac set impuesto = " & DBSet(Rs!importe3, "N") & " where letraser = " & DBSet(Rs!nombre1, "T")
        Sql2 = Sql2 & " and numfactu = " & DBSet(Rs!Importe1, "N") & " and fecfactu = " & DBSet(Rs!Fecha1, "F")
        
        Conn.Execute Sql2
        
        Rs.MoveNext
    Wend
    
    Set Rs = Nothing
    
    CorregirImpuesto = True
    
eCorregirImpuesto:
    If Err.Number <> 0 Then
        MsgBox "Error modificando impuestos " & Err.Description, vbExclamation
    End If
End Function


Private Sub CargarProgres(ByRef PBar As ProgressBar, Valor As Long)
On Error Resume Next
    PBar.Max = 100
    PBar.Value = 0
    PBar.Tag = Valor
    If Err.Number <> 0 Then Err.Clear
End Sub

Private Sub IncrementarProgres(ByRef PBar As ProgressBar, Veces As Integer)
On Error Resume Next
    PBar.Value = PBar.Value + ((Veces * PBar.Max) / CLng(PBar.Tag))
    If Err.Number <> 0 Then Err.Clear
End Sub

