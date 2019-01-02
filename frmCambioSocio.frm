VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmCambioSocio 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cambio de Cliente"
   ClientHeight    =   3915
   ClientLeft      =   45
   ClientTop       =   2430
   ClientWidth     =   7185
   Icon            =   "frmCambioSocio.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3915
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
      Height          =   3630
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   6915
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   4875
         TabIndex        =   3
         Top             =   2595
         Width           =   975
      End
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   3690
         TabIndex        =   2
         Top             =   2595
         Width           =   975
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   0
         Left            =   1845
         MaxLength       =   6
         TabIndex        =   0
         Top             =   840
         Width           =   830
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   1
         Left            =   1845
         MaxLength       =   6
         TabIndex        =   1
         Top             =   1800
         Width           =   830
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   2745
         Locked          =   -1  'True
         TabIndex        =   5
         Text            =   "Text5"
         Top             =   840
         Width           =   3135
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Cliente Destino"
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
         Index           =   0
         Left            =   405
         TabIndex        =   7
         Top             =   1800
         Width           =   1080
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Cliente Origen"
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
         Left            =   405
         TabIndex        =   6
         Top             =   630
         Width           =   1020
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   0
         Left            =   1560
         MouseIcon       =   "frmCambioSocio.frx":000C
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar cliente"
         Top             =   840
         Width           =   240
      End
   End
End
Attribute VB_Name = "frmCambioSocio"
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
Dim i As Byte
Dim Ajena As String
Dim cad As String

    InicializarVbles
    
    If Not DatosOK Then Exit Sub
       
    
    If CambioDeSocio Then
        MsgBox "Proceso realizado correctamente.", vbExclamation
        Unload Me
    End If
    
    
End Sub

Private Function CambioDeSocio() As Boolean
Dim Sql As String
Dim NuevaCuenta As String
Dim Iban As String

    CambioDeSocio = False


    Conn.BeginTrans
    ConnConta.BeginTrans

    NuevaCuenta = "43." & txtCodigo(1)
    
    'Rellenamos si procede
    NuevaCuenta = RellenaCodigoCuenta(NuevaCuenta)

    '==========
    If Not EsCuentaUltimoNivel(NuevaCuenta) Then
        devuelve = "No es cuenta de último nivel: " & Cuenta
        
        Conn.RollbackTrans
        ConnConta.RollbackTrans

        Exit Function
    End If
    '==================
    Dim vSocio As CSocio
    If vSocio.LeerDatos(txtCodigo(0)) Then
    
        Sql = "update ssocio set codsocio = " & DBSet(txtCodigo(1), "N") & ", codmacta = " & DBSet(NuevaCuenta, "T") & " where codsocio = " & DBSet(txtCodigo(0).Text, "N")
        Conn.Execute Sql
        
        Iban = vSocio.Iban & Format(vSocio.Banco, "0000") & Format(vSocio.Sucursal, "0000") & Format(vSocio.Digcontrol, "00") & Right("0000000000" & vSocio.CuentaBan, 10)
        
        If DevuelveDesdeBDNew(cConta, "cuentas", "codmacta", "codmacta", NuevaCuenta, "T") <> "" Then
'            Sql = "update cuentas set nommacta = " & DBSet(vSocio.Nombre, "T") & ", apudirec = 'S', razosoci = " & DBSet(vSocio.Nombre, "T") & ", dirdatos = "
'            Sql = Sql & DBSet(vSocio.Domicilio, "T") & ", codposta = " & DBSet(vSocio.CPostal, "T") & ", despobla = " & DBSet(vSocio.POBLACION, "T")
'            Sql = Sql & ", desprovi = " & DBSet(vSocio.Provincia, "T") & ",nifdatos = " & DBSet(vSocio.NIF, "T") & ", iban = " & DBSet(Iban, "T")
'            Sql = Sql & " where codmacta = " & DBSet(NuevaCuenta, "T")
'
'            ConnConta.Execute Sql

            MsgBox "La cuenta " & NuevaCuenta & " existe en la contabilidad. Revise.", vbExclamation
            
            Conn.RollbackTrans
            ConnConta.RollbackTrans
            
            Exit Function
    

        Else
            Sql = "insert into cuentas (codmacta,nommacta,apudirec,razosoci,dirdatos,codposta,despobla,desprovi,nifdatos,iban) values ("
            Sql = Sql & DBSet(NuevaCuenta, "T") & "," & DBSet(vSocio.Nombre, "T") & ",'S'," & DBSet(vSocio.Nombre, "T") & "," & DBSet(vSocio.Domicilio, "T") & ","
            Sql = Sql & DBSet(vSocio.CPostal, "T") & "," & DBSet(vSocio.POBLACION, "T") & "," & DBSet(vSocio.Provincia, "T") & "," & DBSet(vSocio.NIF, "T") & ","
            Sql = Sql & DBSet(Iban, "T")
            
            ConnConta.Execute Sql
        End If
    Else
        Conn.RollbackTrans
        ConnConta.RollbackTrans
        
        Exit Function
    End If

    Set vSocio = Nothing
    
    CambioDeSocio = True
    
    Conn.CommitTrans
    ConnConta.CommitTrans
    Exit Function

eCambioDeSocio:
    MuestraError Err.Number, "Cambio de Socio", Err.Description
    
    Conn.RollbackTrans
    ConnConta.RollbackTrans
End Function



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

    'IMAGES para busqueda
     Me.imgBuscar(0).Picture = frmPpal.imgListImages16.ListImages(1).Picture
     Me.imgBuscar(1).Picture = frmPpal.imgListImages16.ListImages(1).Picture

    '###Descomentar
'    CommitConexion
         
    FrameCobrosVisible True, h, w
    indFrame = 5
    tabla = "schfac"
    
    'Esto se consigue poneinedo el cancel en el opcion k corresponda
    Me.cmdCancel.Cancel = True
    Me.Width = w + 70
    Me.Height = h + 350
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

Private Sub frmCol_DatoSeleccionado(CadenaSeleccion As String)
'Form de Consulta de Colectivos
    txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "00")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
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
        Case 0, 1 'CLIENTE
            AbrirFrmClientes (Index)
        
        Case 4 'COLECTIVO
            AbrirFrmColectivo (Index)
        
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
            
        Case 0, 1 'CLIENTE
            txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), "ssocio", "nomsocio", "codsocio", "N")
            If txtCodigo(Index).Text <> "" Then txtCodigo(Index).Text = Format(txtCodigo(Index).Text, "000000")
    End Select
End Sub

Private Sub FrameCobrosVisible(visible As Boolean, ByRef h As Integer, ByRef w As Integer)
    Me.FrameCobros.visible = visible
    If visible = True Then
        Me.FrameCobros.Top = -90
        Me.FrameCobros.Left = 0
        Me.FrameCobros.Height = 3630
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

Private Sub AbrirFrmClientes(indice As Integer)
    indCodigo = indice
    Set frmcli = New frmManClien
    frmcli.DatosADevolverBusqueda = "0|1|"
    frmcli.DeConsulta = True
    frmcli.CodigoActual = txtCodigo(indCodigo)
    frmcli.Show vbModal
    Set frmcli = Nothing
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



Private Function DatosOK() As Boolean
Dim b As Boolean

    DatosOK = False
    b = True
    If txtCodigo(0).Text = "" Then
        MsgBox "Debe introducir el cliente origen.", vbExclamation
        PonerFoco txtCodigo(0)
        b = False
    Else
        If TotalRegistros("select * from ssocio where codsocio = " & DBSet(txtCodigo(0).Text, "N")) = 0 Then
            MsgBox "El cliente origen no existe. Reintroduzca.", vbExclamation
            PonerFoco txtCodigo(0)
            b = False
        End If
    End If
    
    If b Then
        If txtcodig(1).Text = "" Then
            MsgBox "Debe introducir el cliente destino.", vbExclamation
            PonerFoco txtCodigo(1)
            b = False
        Else
            If TotalRegistros("select * from ssocio where codsocio = " & DBSet(txtCodigo(1).Text, "N")) > 0 Then
                MsgBox "El cliente destino existe. Reintroduzca.", vbExclamation
                PonerFoco txtCodigo(1)
                b = False
            End If
        End If
    End If
    
End Function
