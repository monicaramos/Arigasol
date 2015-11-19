VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmCuadreDiario 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Informe de Cuadre Diario por Turnos"
   ClientHeight    =   2475
   ClientLeft      =   45
   ClientTop       =   2430
   ClientWidth     =   6825
   Icon            =   "frmCuadreDiario.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2475
   ScaleWidth      =   6825
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
      Height          =   2175
      Left            =   150
      TabIndex        =   4
      Top             =   120
      Width           =   6555
      Begin VB.ComboBox Combo1 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "frmCuadreDiario.frx":000C
         Left            =   1740
         List            =   "frmCuadreDiario.frx":000E
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Tag             =   "Tipo Familia|N|N|0|9|sfamia|tipfamia|||"
         Top             =   930
         Width           =   1050
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   0
         Left            =   1740
         MaxLength       =   10
         TabIndex        =   0
         Tag             =   "Código Postal|T|S|||clientes|codposta|||"
         Top             =   510
         Width           =   1050
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   4935
         TabIndex        =   3
         Top             =   1440
         Width           =   975
      End
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   3750
         TabIndex        =   2
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "Turno"
         ForeColor       =   &H00972E0B&
         Height          =   255
         Index           =   0
         Left            =   600
         TabIndex        =   6
         Top             =   930
         Width           =   765
      End
      Begin VB.Label Label4 
         Caption         =   "Fecha"
         ForeColor       =   &H00972E0B&
         Height          =   255
         Index           =   16
         Left            =   600
         TabIndex        =   5
         Top             =   510
         Width           =   765
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   0
         Left            =   1380
         Picture         =   "frmCuadreDiario.frx":0010
         ToolTipText     =   "Buscar fecha"
         Top             =   510
         Width           =   240
      End
   End
End
Attribute VB_Name = "frmCuadreDiario"
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

Private WithEvents frmFPa As frmManFpago 'F.Pago
Attribute frmFPa.VB_VarHelpID = -1
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
InicializarVbles
    
    '========= PARAMETROS  =============================
    'Añadir el parametro de Empresa
    cadParam = cadParam & "|pEmpresa=""" & vEmpresa.nomEmpre & """|"
    numParam = numParam + 1
    cadParam = cadParam & "|pUsu=" & vSesion.Codigo & "|"
    numParam = numParam + 1
    
    'Añadir el parametro de la fecha
    cadParam = cadParam & "pFecha=""" & Format(txtCodigo(0).Text, "dd/mm/yyyy") & """|"
    numParam = numParam + 1
    
    'Añadir el parametro del turno
    cadParam = cadParam & "pTurno=" & Combo1.ListIndex & "|"
    numParam = numParam + 1
    
    'Comprobar si hay registros a Mostrar antes de abrir el Informe
    cadTABLA = tabla
    
    cadSelect = "fecalbar = " & DBSet(txtCodigo(0).Text, "F")
    If Combo1.ListIndex <> 0 Then
        cadSelect = cadSelect & " and " & "(codturno = " & Combo1.ListIndex & ")"
    End If
    
    'Fecha de albaran
    Codigo = "{scaalb.fecalbar}"
    cDesde = Codigo & " = Date(""" & txtCodigo(0).Text & """) "
    cDesde = "(" & cDesde & ")"

    AnyadirAFormula cadFormula, cDesde

'    'Solo mostramos los carburantes
'    Codigo = "{sfamia.tipfamia}"
'    cDesde = Codigo & " = 1 "
'    cDesde = "(" & cDesde & ")"
'
'    AnyadirAFormula cadFormula, cDesde

    ' mostramos un turno en concreto o todos
    If Combo1.ListIndex <> 0 Then
        Codigo = "{scaalb.codturno}"
        cDesde = Codigo & " = " & Combo1.ListIndex
        AnyadirAFormula cadFormula, cDesde
    End If
    
    
'    cadFormula = "{scaalb.fecalbar} = Date(""" & txtCodigo(0).Text & """) and "
'    cadFormula = cadFormula & "{sfamia.tipfamia} = 1 and ("
'    cadFormula = cadFormula & "{scaalb.codturno} = " & Combo1.ListIndex & " or "
'    cadFormula = cadFormula & Combo1.ListIndex & " = 0 )"
    
    
    If HayRegParaInforme(cadTABLA, cadSelect) Then
          BorradoTablaIntermedia
          
          CargarTablaIntermedia "scaalb"
          CargarTablaIntermedia "sturno"
          cadTitulo = "Informe de Cuadre Diario por Turnos"
'          cadNombreRPT = "rCuadreDiario2.rpt"
          
          If vParamAplic.Cooperativa = 4 Then
            cadParam = cadParam & "pImporteTurno=" & TransformaComasPuntos(ImporteSinFormato(DevuelveValor("select sum(importe3) from tmpinformes where codusu = " & vSesion.Codigo))) & "|"
            numParam = numParam + 1
          End If
          
          
          ' ### [Monica] 15/03/2007
          '****************************
          Dim indRPT As Byte 'Indica el tipo de Documento en la tabla "scryst"
          Dim nomDocu As String 'Nombre de Informe rpt de crystal
         
          indRPT = 2 'Cuadre Diario
         
          If Not PonerParamRPT(indRPT, cadParam, numParam, nomDocu) Then Exit Sub
          'Nombre fichero .rpt a Imprimir
          frmImprimir.NombreRPT = nomDocu
          cadNombreRPT = nomDocu  '  "rCuadreDiario2.rpt"
          '****************************
          
          LlamarImprimir
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
    limpiar Me

    '###Descomentar
'    CommitConexion
         
    FrameCobrosVisible True, h, w
    indFrame = 5
    tabla = "scaalb"
            
    'Esto se consigue poneinedo el cancel en el opcion k corresponda
    Me.cmdCancel.Cancel = True
    
    CargaCombo
    
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

Private Sub CargaCombo()
Dim cad As String
Dim RS As ADODB.Recordset
Dim i As Integer

    On Error GoTo ErrCarga
    Combo1.Clear
    
    'cargamos el combo del turno
    If vParamAplic.Cooperativa <> 4 Then ' Para la pobla solo dejamos meter un turno
        Combo1.AddItem "Todos"
        Combo1.ItemData(Combo1.NewIndex) = 0
    Else
        Combo1.AddItem ""
        Combo1.ItemData(Combo1.NewIndex) = 0
    End If
    
    
    For i = 1 To 9
        Combo1.AddItem i
        Combo1.ItemData(Combo1.NewIndex) = i
    Next i
    
    Exit Sub
    
ErrCarga:
    If Err.Number <> 0 Then MuestraError Err.Number, "Cargar datos combo.", Err.Description
End Sub


Private Sub CargarTablaIntermedia(tabla As String)
Dim RS As ADODB.Recordset
Dim Rs1 As ADODB.Recordset
Dim Sql As String
Dim sql2 As String
Dim Sql3 As String
Dim Fam As String
Dim NumLin As Integer

'    CargarPb1 (tabla)
    
    If tabla = "scaalb" Then ' slhfac
        Sql = "select sartic.codartic, sum(cantidad), sum(importel) from scaalb, sartic, sfamia  where "
        Sql = Sql & " scaalb.fecalbar = " & DBSet(txtCodigo(0), "F")
        If Combo1.ListIndex <> 0 Then
            Sql = Sql & " and (scaalb.codturno = " & Combo1.ListIndex & ")"
        End If
        Sql = Sql & " and sfamia.tipfamia = 1 and sfamia.codfamia = sartic.codfamia and sartic.codartic = scaalb.codartic"
    Else
        If vParamAplic.Cooperativa = 4 Then
            Sql = "select sartic.codartic, round(sum(contafin - containi) * preventa,2), 0 from sturno, sartic, sfamia  where "
            Sql = Sql & " sturno.fechatur = " & DBSet(txtCodigo(0), "F")
            If Combo1.ListIndex <> 0 Then
                Sql = Sql & " and (sturno.codturno = " & Combo1.ListIndex & ")"
            End If
            Sql = Sql & " and sfamia.tipfamia = 1 and sfamia.codfamia = sartic.codfamia and sartic.codartic = sturno.codartic"
        Else
            Sql = "select sartic.codartic, sum(litrosve), sum(importel) from sturno, sartic, sfamia  where "
            Sql = Sql & " sturno.fechatur = " & DBSet(txtCodigo(0), "F")
            If Combo1.ListIndex <> 0 Then
                Sql = Sql & " and (sturno.codturno = " & Combo1.ListIndex & ")"
            End If
            Sql = Sql & " and sfamia.tipfamia = 1 and sfamia.codfamia = sartic.codfamia and sartic.codartic = sturno.codartic"
        End If
    End If
    Sql = Sql & " group by 1"

        
    Set RS = New ADODB.Recordset
    RS.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    While Not RS.EOF
        Sql = ""
        Sql = DevuelveDesdeBDNew(cPTours, "tmpinformes", "codigo1", "codigo1", RS!codArtic, "N", , "codusu", vSesion.Codigo, "N")
        If Sql = "" Then 'insertamos
            Sql = "insert into tmpinformes (codusu,codigo1, importe1, importe2, importe3, importe4) values ("
            Sql = Sql & DBSet(vSesion.Codigo, "N") & "," & DBSet(RS!codArtic, "N") & ","
            
            If tabla = "scaalb" Then
                Sql = Sql & DBSet(RS.Fields(1).Value, "N") & "," & DBSet(RS.Fields(2).Value, "N") & ",0,0)"
            Else
                Sql = Sql & "0,0," & DBSet(RS.Fields(1).Value, "N") & "," & DBSet(RS.Fields(2).Value, "N") & ")"
            End If
        Else 'actualizamos
            Sql = "update tmpinformes set "
            If tabla = "scaalb" Then
                Sql = Sql & "importe1 = importe1 + " & DBSet(RS.Fields(1).Value, "N") & ", "
                Sql = Sql & "importe2 = importe2 + " & DBSet(RS.Fields(2).Value, "N")
                Sql = Sql & " where codusu = " & vSesion.Codigo & " and codigo1 = " & DBSet(RS!codArtic, "N")
                
            Else
                Sql = Sql & "importe3 = importe3 + " & DBSet(RS.Fields(1).Value, "N") & ", "
                Sql = Sql & "importe4 = importe4 + " & DBSet(RS.Fields(2).Value, "N")
                Sql = Sql & " where codusu = " & vSesion.Codigo & " and codigo1 = " & DBSet(RS!codArtic, "N")
            End If
        End If
        Conn.Execute Sql
        RS.MoveNext
    Wend
    
    Sql = "update tmpinformes set fecha1 = " & DBSet(txtCodigo(0).Text, "F") & ",campo1 = " & DBSet(Combo1.ListIndex, "N")
    Sql = Sql & " where codusu = " & vSesion.Codigo
    
    Conn.Execute Sql
    RS.Close
    Set RS = Nothing
End Sub


Private Sub BorradoTablaIntermedia()
Dim Sql As String

    Sql = "delete from tmpinformes where codusu = " & vSesion.Codigo
    Conn.Execute Sql
    
End Sub

