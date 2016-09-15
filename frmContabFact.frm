VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmContabFact 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Contabilización de Facturas "
   ClientHeight    =   5955
   ClientLeft      =   45
   ClientTop       =   2430
   ClientWidth     =   6645
   Icon            =   "frmContabFact.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5955
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
      Height          =   5535
      Left            =   90
      TabIndex        =   10
      Top             =   120
      Width           =   6375
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   8
         Left            =   1590
         MaxLength       =   6
         TabIndex        =   1
         Tag             =   "Código Propio|N|N|1|99|sbanco|codbanpr|00|S|"
         Top             =   1170
         Width           =   830
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   8
         Left            =   2490
         Locked          =   -1  'True
         TabIndex        =   21
         Text            =   "Text5"
         Top             =   1170
         Width           =   3135
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   7
         Left            =   1575
         MaxLength       =   10
         TabIndex        =   0
         Tag             =   "Código Postal|T|S|||clientes|codposta|||"
         Top             =   630
         Width           =   1050
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   5
         Left            =   3720
         MaxLength       =   3
         TabIndex        =   3
         Top             =   1995
         Width           =   495
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   4
         Left            =   1605
         MaxLength       =   3
         TabIndex        =   2
         Top             =   1980
         Width           =   495
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   3
         Left            =   3735
         MaxLength       =   10
         TabIndex        =   7
         Tag             =   "Código Postal|T|S|||clientes|codposta|||"
         Top             =   3210
         Width           =   1050
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   2
         Left            =   1605
         MaxLength       =   10
         TabIndex        =   6
         Tag             =   "Código Postal|T|S|||clientes|codposta|||"
         Top             =   3180
         Width           =   1050
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   4755
         TabIndex        =   9
         Top             =   4980
         Width           =   975
      End
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   3570
         TabIndex        =   8
         Top             =   4980
         Width           =   975
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   0
         Left            =   1605
         MaxLength       =   7
         TabIndex        =   4
         Tag             =   "Nº de Factura|N|N|0|9999999|schfac|numfactu|0000000|S|"
         Top             =   2580
         Width           =   830
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   1
         Left            =   3720
         MaxLength       =   7
         TabIndex        =   5
         Tag             =   "Nº de Factura|N|N|0|9999999|schfac|numfactu|0000000|S|"
         Top             =   2610
         Width           =   830
      End
      Begin MSComctlLib.ProgressBar Pb1 
         Height          =   255
         Left            =   330
         TabIndex        =   20
         Top             =   4380
         Width           =   5430
         _ExtentX        =   9578
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Image imgAyuda 
         Height          =   240
         Index           =   0
         Left            =   2700
         MousePointer    =   4  'Icon
         Tag             =   "-1"
         ToolTipText     =   "Ayuda"
         Top             =   630
         Width           =   240
      End
      Begin VB.Label lblProgres 
         Height          =   285
         Index           =   1
         Left            =   330
         TabIndex        =   25
         Top             =   3990
         Width           =   5295
      End
      Begin VB.Label lblProgres 
         Height          =   285
         Index           =   0
         Left            =   330
         TabIndex        =   24
         Top             =   3720
         Width           =   5265
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Banco Propio"
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
         Index           =   5
         Left            =   330
         TabIndex        =   23
         Top             =   960
         Width           =   930
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   8
         Left            =   1290
         MouseIcon       =   "frmContabFact.frx":000C
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar cliente"
         Top             =   1170
         Width           =   240
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   7
         Left            =   1260
         Picture         =   "frmContabFact.frx":015E
         ToolTipText     =   "Buscar fecha"
         Top             =   630
         Width           =   240
      End
      Begin VB.Label Label4 
         Caption         =   "Fecha Vto"
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
         Index           =   4
         Left            =   330
         TabIndex        =   22
         Top             =   450
         Width           =   1815
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
         TabIndex        =   19
         Top             =   1740
         Width           =   1005
      End
      Begin VB.Label Label4 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   1
         Left            =   2880
         TabIndex        =   18
         Top             =   1995
         Width           =   420
      End
      Begin VB.Label Label4 
         Caption         =   "Desde"
         Height          =   195
         Index           =   0
         Left            =   720
         TabIndex        =   17
         Top             =   1980
         Width           =   465
      End
      Begin VB.Label Label4 
         Caption         =   "Fecha Factura"
         ForeColor       =   &H00972E0B&
         Height          =   255
         Index           =   16
         Left            =   330
         TabIndex        =   16
         Top             =   2940
         Width           =   1815
      End
      Begin VB.Label Label4 
         Caption         =   "Desde"
         Height          =   195
         Index           =   15
         Left            =   720
         TabIndex        =   15
         Top             =   3180
         Width           =   465
      End
      Begin VB.Label Label4 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   14
         Left            =   2880
         TabIndex        =   14
         Top             =   3210
         Width           =   420
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   2
         Left            =   1290
         Picture         =   "frmContabFact.frx":01E9
         ToolTipText     =   "Buscar fecha"
         Top             =   3180
         Width           =   240
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   3
         Left            =   3420
         Picture         =   "frmContabFact.frx":0274
         ToolTipText     =   "Buscar fecha"
         Top             =   3210
         Width           =   240
      End
      Begin VB.Label Label4 
         Caption         =   "Desde"
         Height          =   195
         Index           =   13
         Left            =   720
         TabIndex        =   13
         Top             =   2580
         Width           =   465
      End
      Begin VB.Label Label4 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   12
         Left            =   2880
         TabIndex        =   12
         Top             =   2625
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
         Left            =   330
         TabIndex        =   11
         Top             =   2340
         Width           =   555
      End
   End
End
Attribute VB_Name = "frmContabFact"
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
Dim cadTabla As String, cOrden As String
Dim cadMen As String
Dim i As Byte
Dim SQL As String
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
    
    ContabilizarFacturas tabla, cadSelect
     'Eliminar la tabla TMP
    BorrarTMPFacturas
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
    limpiar Me

    'IMAGES para busqueda
     Me.imgBuscar(8).Picture = frmPpal.imgListImages16.ListImages(1).Picture
     
    '###Descomentar
'    CommitConexion
         
    FrameCobrosVisible True, h, w
    indFrame = 5
    tabla = "schfac"
    
    Pb1.visible = False
    
    imgAyuda(0).Picture = frmPpal.ImageListB.ListImages(10).Picture
    
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

Private Sub frmBpr_DatoSeleccionado(CadenaSeleccion As String)
'Form de Consulta de Colectivos
    txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "00")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmC_Selec(vFecha As Date)
 'Fecha
    txtCodigo(CByte(imgFec(2).Tag)).Text = Format(vFecha, "dd/MM/yyyy")
End Sub

Private Sub imgAyuda_Click(Index As Integer)
Dim vCadena As String
    
    Select Case Index
        Case 0
           ' "____________________________________________________________"
            vCadena = "Si ponemos la Fecha de Vencimiento los efectos se calculan sobre" & vbCrLf & _
                      "esta fecha. " & vbCrLf & vbCrLf & _
                      "En caso contrario, el cálculo se realizará sobre la fecha de factura." & vbCrLf & vbCrLf

    End Select
    MsgBox vCadena, vbInformation, "Descripción de Ayuda"

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
        
        Case 8 ' BANCO PROPIO
            AbrirFrmBancoPropio (Index)
        
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
            Case 8: KEYBusqueda KeyAscii, 8 'banco propio
            Case 2: KEYFecha KeyAscii, 2 'fecha desde
            Case 3: KEYFecha KeyAscii, 3 'fecha hasta
            Case 7: KEYFecha KeyAscii, 7 'fecha de vencimiento
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
        Case 8 ' BANCO PROPIO
            If txtCodigo(Index).Text <> "" Then
                PonerFormatoEntero txtCodigo(Index)
                txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), "sbanco", "nombanco", "codbanpr", "N")
                If txtNombre(Index).Text = "" Then
                    MsgBox "El Banco introducido no existe. Reintroduzca.", vbExclamation
                    PonerFoco txtCodigo(Index)
                End If
            End If
            
        Case 2, 3, 7  'FECHAS
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
        Me.FrameCobros.Height = 6015
        Me.FrameCobros.Width = 6555
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

    If txtCodigo(7).Text = "" Then
        If MsgBox("Si no hay Fecha de Vencimiento se utilizará la fecha de factura." & vbCrLf & vbCrLf & "¿Desea continuar?.", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
            b = False
            PonerFoco txtCodigo(7)
        End If
    End If
    If txtCodigo(8).Text = "" Then
        MsgBox "Debe introducir obligatoriamente un Banco para realizar el cobro.", vbExclamation
        b = False
        PonerFoco txtCodigo(8)
    End If

     
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
Private Sub ContabilizarFacturas(cadTabla As String, cadWhere As String)
'Contabiliza Facturas de Clientes o de Proveedores
Dim SQL As String
Dim b As Boolean
Dim tmpErrores As Boolean 'Indica si se creo correctamente la tabla de errores
Dim CCoste As String

    If cadTabla = "schfac" Then
        SQL = "VENCON" 'contabilizar facturas de venta
    End If

    'Bloquear para que nadie mas pueda contabilizar
    DesBloqueoManual (SQL)
    If Not BloqueoManual(SQL, "1") Then
        MsgBox "No se pueden Contabilizar Facturas. Hay otro usuario contabilizando.", vbExclamation
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

    'comprobar si existen en Arigasol facturas anteriores al periodo solicitado
    'sin contabilizar
    If Me.txtCodigo(2).Text <> "" Then
        SQL = "SELECT COUNT(*) FROM " & cadTabla
        If cadTabla = "schfac" Then
            SQL = SQL & " WHERE fecfactu <"
        End If
        SQL = SQL & DBSet(txtCodigo(2), "F") & " AND intconta=0 "
        If RegistrosAListar(SQL) > 0 Then
            MsgBox "Hay Facturas anteriores sin contabilizar.", vbExclamation
            Exit Sub
        End If
    End If
    
    'Visualizar la barra de Progreso
    Me.Pb1.visible = True
'    Me.Pb1.Top = 3350
    
    
    '==========================================================
    'REALIZAR COMPROBACIONES ANTES DE CONTABILIZAR FACTURAS
    '==========================================================
    
    Me.lblProgres(0).Caption = "Comprobaciones: "
    CargarProgres Me.Pb1, 100
        
    BorrarTMPFacturas
    'Cargar tabla TEMP con las Facturas que vamos a Trabajar
    b = CrearTMPFacturas(cadTabla, cadWhere)
    If Not b Then Exit Sub
            
    ' nuevo
    b = CrearTMPErrComprob()
    If Not b Then Exit Sub
    
    'comprobar que todas las LETRAS SERIE existen en la contabilidad y en Arigasol
    '-----------------------------------------------------------------------------
    IncrementarProgres Me.Pb1, 10
    If cadTabla = "schfac" Then
        Me.lblProgres(1).Caption = "Comprobando letras de serie ..."
        b = ComprobarLetraSerie(cadTabla)
    End If
    IncrementarProgres Me.Pb1, 10
    Me.Refresh
    If Not b Then
        frmMensaje.OpcionMensaje = 1
        frmMensaje.Show vbModal
        Exit Sub
    End If
    
    'comprobar que no haya Nº FACTURAS en la contabilidad para esa fecha
    'que ya existan
    '-----------------------------------------------------------------------
    If cadTabla = "schfac" Then
        Me.lblProgres(1).Caption = "Comprobando Nº Facturas en contabilidad ..."
        If vParamAplic.ContabilidadNueva Then
            SQL = "anofactu>=" & Year(txtCodigo(2).Text) & " AND anofactu<= " & Year(txtCodigo(3).Text)
        Else
            SQL = "anofaccl>=" & Year(txtCodigo(2).Text) & " AND anofaccl<= " & Year(txtCodigo(3).Text)
        End If
        b = ComprobarNumFacturas(cadTabla, SQL)
    End If
    IncrementarProgres Me.Pb1, 20
    Me.Refresh
    If Not b Then
        frmMensaje.OpcionMensaje = 1
        frmMensaje.Show vbModal
        Exit Sub
    End If
    
    'comprobar que todas las CUENTAS de los distintos clientes que vamos a
    'contabilizar existen en la Conta: sclien.codmacta IN (conta.cuentas.codmacta)
    '-----------------------------------------------------------------------------
    Me.lblProgres(1).Caption = "Comprobando Cuentas Contables en contabilidad ..."
    b = ComprobarCtaContable(cadTabla, 1)
    IncrementarProgres Me.Pb1, 20
    Me.Refresh
    If Not b Then
        frmMensaje.OpcionMensaje = 2
        frmMensaje.Show vbModal
        Exit Sub
    End If
    
    'comprobar que todas las CUENTAS de venta de la familia de los articulos que vamos a
    'contabilizar existen en la Conta: sfamia.ctaventa IN (conta.cuentas.codmacta)
    '-----------------------------------------------------------------------------
    If cadTabla = "schfac" Then
        Me.lblProgres(1).Caption = "Comprobando Cuentas Ctbles Ventas en contabilidad ..."
    End If
    b = ComprobarCtaContable(cadTabla, 2)
    IncrementarProgres Me.Pb1, 20
    Me.Refresh
    If Not b Then
        frmMensaje.OpcionMensaje = 2
        frmMensaje.Show vbModal
        Exit Sub
    End If
    
    'comprobar que todas las CUENTAS de venta de la familia de los articulos que vamos a
    'contabilizar son de grupo de ventas: empiezan por conta.parametros.grupovtas
    '-----------------------------------------------------------------------------
    If cadTabla = "schfac" Then
        Me.lblProgres(1).Caption = "Comprobando Cuentas Ctbles Ventas en contabilidad ..."
    End If
    b = ComprobarCtaContable(cadTabla, 3)
    IncrementarProgres Me.Pb1, 20
    Me.Refresh
    If Not b Then
        frmMensaje.OpcionMensaje = 2
        frmMensaje.Show vbModal
        Exit Sub
    End If
    
    
    
    
    'comprobar que todas la CUENTA del banco propio donde contabilizar el cobro
    'que existen en la Conta: sbanpr.codmacta IN (conta.cuentas.codmacta)
    '-----------------------------------------------------------------------------
    Me.lblProgres(1).Caption = "Comprobando Cuentas Contables del Banco en contabilidad ..."
    
    b = ComprobarCtaContable(CStr(txtCodigo(8).Text), 4)
    IncrementarProgres Me.Pb1, 20
    Me.Refresh
    If Not b Then
        frmMensaje.OpcionMensaje = 2
        frmMensaje.Show vbModal
        Exit Sub
    End If
    
    
    'comprobar que todos las TIPO IVA de las distintas facturas que vamos a
    'contabilizar existen en la Conta: schfac.codigiv1,codigiv2,codigiv3 IN (conta.tiposiva.codigiva)
    '--------------------------------------------------------------------------
    Me.lblProgres(1).Caption = "Comprobando Tipos de IVA en contabilidad ..."
    b = ComprobarTiposIVA(cadTabla)
    IncrementarProgres Me.Pb1, 10
    Me.Refresh
    If Not b Then
        frmMensaje.OpcionMensaje = 3
        frmMensaje.Show vbModal
        Exit Sub
    End If
    
    
    '===========================================================================
    'CONTABILIZAR FACTURAS
    '===========================================================================
    Me.lblProgres(0).Caption = "Contabilizar Facturas: "
    CargarProgres Me.Pb1, 10
    Me.lblProgres(1).Caption = "Insertando Facturas en Contabilidad..."
       
    
    'Crear tabla TEMP para los posible errores de facturas
    tmpErrores = CrearTMPErrFact(cadTabla)
    
    
    b = PasarFacturasAContab(cadTabla, txtCodigo(7).Text, txtCodigo(8).Text, CCoste)
    
    If Not b Then
        If tmpErrores Then
            'Cargar un listview con la tabla TEMP de Errores y mostrar
            'las facturas que fallaron
            frmMensaje.OpcionMensaje = 10
            frmMensaje.Show vbModal
        Else
            MsgBox "No pueden mostrarse los errores.", vbInformation
        End If
    Else
        MsgBox "El proceso ha finalizado correctamente.", vbInformation
    End If
    
    'Eliminar tabla TEMP de Errores
    BorrarTMPErrFact
    BorrarTMPErrComprob
End Sub


Private Function PasarFacturasAContab(cadTabla As String, FecVenci As String, Banpr As String, CCoste As String) As Boolean
Dim SQL As String
Dim Rs As ADODB.Recordset
Dim b As Boolean
Dim i As Integer
Dim numfactu As Integer
Dim codigo1 As String

    On Error GoTo EPasarFac

    PasarFacturasAContab = False
    
    'Total de Facturas a Insertar en la contabilidad
    SQL = "SELECT count(*) "
    SQL = SQL & " FROM " & cadTabla & " INNER JOIN tmpfactu "
    If cadTabla = "schfac" Then
        codigo1 = "letraser"
    End If
    SQL = SQL & " ON " & cadTabla & "." & codigo1 & "=tmpfactu." & codigo1
    SQL = SQL & " AND " & cadTabla & ".numfactu=tmpfactu.numfactu AND " & cadTabla & ".fecfactu=tmpfactu.fecfactu "
    
    
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not Rs.EOF Then
        numfactu = Rs.Fields(0)
    Else
        numfactu = 0
    End If
    Rs.Close
    Set Rs = Nothing

    If numfactu > 0 Then
        CargarProgres Me.Pb1, numfactu
        
        SQL = "SELECT * "
        SQL = SQL & " FROM tmpfactu "
            
        Set Rs = New ADODB.Recordset
        Rs.Open SQL, Conn, adOpenStatic, adLockPessimistic, adCmdText
        i = 1

        b = True
        'contabilizar cada una de las facturas seleccionadas
        While Not Rs.EOF
            If cadTabla = "schfac" Then
                SQL = cadTabla & "." & codigo1 & "=" & DBSet(Rs.Fields(0), "T") & " and numfactu=" & DBLet(Rs!numfactu, "N")
                SQL = SQL & " and fecfactu=" & DBSet(Rs!fecfactu, "F")
                If PasarFactura(SQL, FecVenci, Banpr, CCoste) = False And b Then b = False
            End If
            
            IncrementarProgres Me.Pb1, 1
            Me.lblProgres(1).Caption = "Insertando Facturas en Contabilidad...   (" & i & " de " & numfactu & ")"
            Me.Refresh
            i = i + 1
            Rs.MoveNext
        Wend
        Rs.Close
        Set Rs = Nothing
    End If
    
EPasarFac:
    If Err.Number <> 0 Then b = False
    
    If b Then
        PasarFacturasAContab = True
    Else
        PasarFacturasAContab = False
    End If
End Function


