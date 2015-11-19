VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmCieTurnoReg 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cierre de Turno"
   ClientHeight    =   2805
   ClientLeft      =   45
   ClientTop       =   2430
   ClientWidth     =   6630
   Icon            =   "frmCieTurnoReg.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2805
   ScaleWidth      =   6630
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
      Height          =   2685
      Left            =   30
      TabIndex        =   4
      Top             =   60
      Width           =   6555
      Begin VB.Frame Frame1 
         Caption         =   "Datos para Selección"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   1365
         Left            =   450
         TabIndex        =   5
         Top             =   300
         Width           =   5595
         Begin VB.TextBox txtcodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   1
            Left            =   1980
            MaxLength       =   1
            TabIndex        =   1
            Top             =   930
            Width           =   330
         End
         Begin VB.TextBox txtcodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   0
            Left            =   1980
            MaxLength       =   10
            TabIndex        =   0
            Tag             =   "Código Postal|T|S|||clientes|codposta|||"
            Top             =   570
            Width           =   1080
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Nº Turno"
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
            TabIndex        =   7
            Top             =   930
            Width           =   645
         End
         Begin VB.Label Label4 
            Caption         =   "Fecha"
            ForeColor       =   &H00972E0B&
            Height          =   255
            Index           =   16
            Left            =   300
            TabIndex        =   6
            Top             =   570
            Width           =   1425
         End
         Begin VB.Image imgFec 
            Height          =   240
            Index           =   0
            Left            =   1680
            Picture         =   "frmCieTurnoReg.frx":000C
            ToolTipText     =   "Buscar fecha"
            Top             =   570
            Width           =   240
         End
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   5085
         TabIndex        =   3
         Top             =   2070
         Width           =   975
      End
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   3900
         TabIndex        =   2
         Top             =   2070
         Width           =   975
      End
   End
End
Attribute VB_Name = "frmCieTurnoReg"
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

Private WithEvents frmFpa As frmManFpago 'F.Pago
Attribute frmFpa.VB_VarHelpID = -1
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
Dim i As Byte
Dim b As Boolean

InicializarVbles
    
    '========= PARAMETROS  =============================
    'Añadir el parametro de Empresa
    cadParam = cadParam & "|pEmpresa=""" & vEmpresa.nomEmpre & """|"
    numParam = numParam + 1
    cadParam = cadParam & "pUsu=" & vSesion.Codigo & "|"
    numParam = numParam + 1
    
    'Añadir el parametro de la fecha
    cadParam = cadParam & "pFecha=""" & Format(txtCodigo(0).Text, "dd/mm/yyyy") & """|"
    numParam = numParam + 1
    
    'Añadir el parametro del turno
    cadParam = cadParam & "pTurno=" & txtCodigo(1).Text & "|"
    numParam = numParam + 1
    
    'Comprobar si hay registros a Mostrar antes de abrir el Informe
    cadTABLA = tabla
    
    cadSelect = "fechatur = " & DBSet(txtCodigo(0).Text, "F") & " and "
    If txtCodigo(1).Text <> "" Then
        cadSelect = cadSelect & "(codturno = " & txtCodigo(1).Text & ")"
    End If
    
    cadFormula = "{sturno.fechatur} = cdate(""" & Format(txtCodigo(0).Text, "dd/mm/yyyy") & """)"
    cadFormula = cadFormula & " and {sturno.codturno} = " & Format(txtCodigo(1).Text, "0")
    If HayRegParaInforme(cadTABLA, cadSelect) Then
          BorradoTablaIntermedia
          b = CargarTablaIntermedia
          If b Then
            cadTitulo = "Cuadre Diario de Turno"
            cadNombreRPT = "rCuadreDiarioReg.rpt"
            LlamarImprimir
          End If
          'AbrirVisReport
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub


Private Sub Combo1_KeyDown(KeyCode As Integer, Shift As Integer)
    KEYdown KeyCode
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
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

    '###Descomentar
'    CommitConexion
         
    FrameCobrosVisible True, h, w
    indFrame = 5
    tabla = "sturno"
            
    'Esto se consigue poneinedo el cancel en el opcion k corresponda
    Me.cmdCancel.Cancel = True
    
    
'   Me.Width = w + 70
'   Me.Height = h + 350
End Sub

Private Sub frmC_Selec(vFecha As Date)
 'Fecha
    txtCodigo(CByte(imgFec(0).Tag)).Text = Format(vFecha, "dd/MM/yyyy")
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
    imgFec(0).Tag = Index 'independentment de les dates que tinga, sempre pose l'index en la 27
    If txtCodigo(Index).Text <> "" Then frmC.NovaData = txtCodigo(Index).Text

    frmC.Show vbModal
    Set frmC = Nothing
'    PonerFoco txtCodigo(CByte(imgFec(0).Tag) + 2)
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
'14/02/2007
'    KEYpress KeyAscii
'ahora
    If KeyAscii = teclaBuscar Then
        Select Case Index
            Case 0: KEYFecha KeyAscii, 0 'fecha de turno
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
        Case 0 'FECHAS
            If txtCodigo(Index).Text <> "" Then PonerFormatoFecha txtCodigo(Index)
            
        Case 1 'turno
            If txtCodigo(Index).Text <> "" Then PonerFormatoEntero txtCodigo(Index)
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
Dim Rs As ADODB.Recordset
Dim Rs2 As ADODB.Recordset
Dim Rs3 As ADODB.Recordset
Dim sql As String
Dim Sql2 As String
Dim sql3 As String
Dim vInvent As Currency
Dim vAcumul As Currency
Dim vCompra As Currency
Dim vStock As Currency
Dim cInvent As String
Dim FecInv As String
Dim TurnoInv As String
Dim LineaInv As String

    On Error GoTo eCargarTablaIntermedia
    
    CargarTablaIntermedia = False
    
    sql = "select numtanqu, litrosve from sturno "
    sql = sql & " where fechatur = " & DBSet(txtCodigo(0).Text, "F") & " and "
    sql = sql & " codturno = " & DBSet(txtCodigo(1).Text, "N") & " and "
    sql = sql & " tiporegi = 1 " 'cogemos solo tanques

        
    Set Rs = New ADODB.Recordset
    Rs.Open sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    While Not Rs.EOF
        '*****
        ' seleccionamos la maxima fecha, en que turno y linea se hizo varillas (inventario) para ese tanque
        Sql2 = "select max(fechatur) from sturno where numtanqu = " & DBLet(Rs.Fields(0).Value, "N")
        Sql2 = Sql2 & " and "
        Sql2 = Sql2 & "((fechatur = " & DBSet(txtCodigo(0).Text, "F") & " and codturno <= " & DBSet(txtCodigo(1).Text, "N") & ") or "
        Sql2 = Sql2 & " (fechatur < " & DBSet(txtCodigo(0).Text, "F") & "))"
        ' end del añadido
        
        Sql2 = Sql2 & " and tiporegi = 4 "
        Set Rs2 = New ADODB.Recordset
        Rs2.Open Sql2, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        FecInv = "01/01/1900"
        TurnoInv = 1
        If Not Rs2.EOF Then
            FecInv = DBLet(Rs2.Fields(0).Value, "F")
            If FecInv = "" Then FecInv = "01/01/1900"
            Set Rs2 = Nothing
            
            Sql2 = "select max(codturno) from sturno where numtanqu = " & DBLet(Rs.Fields(0).Value, "N")
            Sql2 = Sql2 & " and tiporegi = 4 "
            Sql2 = Sql2 & " and fechatur = " & DBSet(FecInv, "F")
            
            Set Rs2 = New ADODB.Recordset
            Rs2.Open Sql2, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            TurnoInv = 0
            If Not Rs2.EOF Then
                TurnoInv = DBLet(Rs2.Fields(0).Value, "N")
                If TurnoInv = 0 Then TurnoInv = 1
                
                Set Rs2 = Nothing
                
                Sql2 = "select max(numlinea) from sturno where numtanqu = " & DBLet(Rs.Fields(0).Value, "N")
                Sql2 = Sql2 & " and tiporegi = 4 "
                Sql2 = Sql2 & " and fechatur = " & DBSet(FecInv, "F")
                Sql2 = Sql2 & " and codturno = " & DBSet(TurnoInv, "N")

                Set Rs2 = New ADODB.Recordset
                Rs2.Open Sql2, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                LineaInv = 0
                If Not Rs2.EOF Then
                    LineaInv = DBLet(Rs2.Fields(0).Value, "N")
                End If
            
            End If
        End If
        
        '*****
        ' obtenemos la cantidad del ultimo registro de inventario
        cInvent = ""
        If Not Rs2.EOF Then
            cInvent = DevuelveDesdeBDNew(cPTours, "sturno", "litrosve", "fechatur", FecInv, "F", , "codturno", TurnoInv, "N", "numlinea", LineaInv, "N")
        End If
        If cInvent = "" Then
            vInvent = 0
        Else
            vInvent = CCur(cInvent)
        End If
        '*****
        ' obtenemos el acumulado que hay entre el inventario y el consumo de fecha y turno
        ' donde :fecinv, turnoinv, lineainv: CP del registro de ultimo inventario
        '        p_fecha y p_turno: fecha y turno que pedimos en el programa
        '
        'select sum(litrosve)
        'From sturno
        'where numtanqu = 1 and tiporegi = 1 and
        '     ((fechatur = $fecinv and codturno = $turnoinv and numlinea > $lineainv) or
        '      (fechatur = $fecinv and codturno > $turnoinv) or
        '      (fechatur > $fecinv)) and
        '     ((fechatur = $p_fecha and codturno < $p_turno) or
        '      (fechatur < $p_fecha))
        sql3 = "select sum(litrosve) from sturno where numtanqu = " & DBLet(Rs.Fields(0).Value, "N") & " and tiporegi = 1 and "
'        sql3 = sql3 & "((fechatur = " & DBSet(FecInv, "F") & " and codturno = " & DBSet(TurnoInv, "N") & " and numlinea > " & DBSet(LineaInv, "N") & ") or "
        sql3 = sql3 & " ((fechatur = " & DBSet(FecInv, "F") & " and codturno > " & DBSet(TurnoInv, "N") & ") or "
        sql3 = sql3 & " (fechatur > " & DBSet(FecInv, "F") & "))    and "
        sql3 = sql3 & "((fechatur = " & DBSet(txtCodigo(0).Text, "F") & " and codturno <= " & DBSet(txtCodigo(1).Text, "N") & ") or "
        sql3 = sql3 & " (fechatur < " & DBSet(txtCodigo(0).Text, "F") & "))"
        Set Rs3 = New ADODB.Recordset
        Rs3.Open sql3, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        vAcumul = 0
        If Not Rs3.EOF Then
            vAcumul = DBLet(Rs3.Fields(0).Value, "N")
        End If
        Set Rs3 = Nothing
        
        '*****
        ' obtenemos lo mismo pero las compras que se han realizado
        ' donde :fecinv, turnoinv, lineainv: CP del registro de ultimo inventario
        '        p_fecha y p_turno: fecha y turno que pedimos en el programa
        '
        'select sum(litrosve)
        'From sturno
        'where numtanqu = 1 and tiporegi = 3 and
        '     ((fechatur = $fecinv and codturno = $turnoinv and numlinea > $lineainv) or
        '      (fechatur = $fecinv and codturno > $turnoinv) or
        '      (fechatur > $fecinv)) and
        '     ((fechatur = $p_fecha and codturno < $p_turno) or
        '      (fechatur < $p_fecha))
        sql3 = "select sum(litrosve) from sturno where numtanqu = " & DBLet(Rs.Fields(0).Value, "N") & " and tiporegi = 3 and "
'        sql3 = sql3 & "((fechatur = " & DBSet(FecInv, "F") & " and codturno = " & DBSet(TurnoInv, "N") & " and numlinea > " & DBSet(LineaInv, "N") & ") or "
        sql3 = sql3 & " ((fechatur = " & DBSet(FecInv, "F") & " and codturno > " & DBSet(TurnoInv, "N") & ") or "
        sql3 = sql3 & " (fechatur > " & DBSet(FecInv, "F") & "))    and "
        sql3 = sql3 & "((fechatur = " & DBSet(txtCodigo(0).Text, "F") & " and codturno <= " & DBSet(txtCodigo(1).Text, "N") & ") or "
        sql3 = sql3 & " (fechatur < " & DBSet(txtCodigo(0).Text, "F") & "))"
        Set Rs3 = New ADODB.Recordset
        Rs3.Open sql3, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        vCompra = 0
        If Not Rs3.EOF Then
            vCompra = DBLet(Rs3.Fields(0).Value, "N")
        End If
        Set Rs3 = Nothing
        
        vStock = vInvent - vAcumul + vCompra
        
        sql = "insert into tmpinformes (codusu, codigo1, importe1, fecha1, "
        sql = sql & "importe2, importe3, importe4) values ("
        sql = sql & DBSet(vSesion.Codigo, "N") & "," & DBSet(Rs.Fields(0).Value, "N") & "," & DBSet(Rs.Fields(1).Value, "N") & "," & DBSet(FecInv, "F") & ","
        sql = sql & DBSet(vInvent, "N") & "," & DBSet(vCompra, "N") & "," & DBSet(vStock, "N") & ")"
            
        Conn.Execute sql
        Rs.MoveNext
    Wend
    
    Rs.Close
    Set Rs = Nothing
    CargarTablaIntermedia = True
    
eCargarTablaIntermedia:
    If Err.Number <> 0 Then
        MsgBox "No se ha podido hacer la carga previa al listado", vbExclamation
    End If
End Function

Private Sub BorradoTablaIntermedia()
Dim sql As String

    sql = "delete from tmpinformes where codusu = " & vSesion.Codigo
    Conn.Execute sql
End Sub
