VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmContCieTurnoReg 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Integración Cierre Contable de Turno"
   ClientHeight    =   5730
   ClientLeft      =   45
   ClientTop       =   2430
   ClientWidth     =   6825
   Icon            =   "frmContCieTurnoReg.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5730
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
      Height          =   5505
      Left            =   150
      TabIndex        =   7
      Top             =   120
      Width           =   6555
      Begin VB.Frame Frame2 
         Caption         =   "Datos para la contabilización"
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
         Height          =   2025
         Left            =   540
         TabIndex        =   10
         Top             =   1575
         Width           =   5625
         Begin VB.TextBox txtNombre 
            BackColor       =   &H80000018&
            Enabled         =   0   'False
            Height          =   285
            Index           =   2
            Left            =   2700
            Locked          =   -1  'True
            TabIndex        =   16
            Top             =   390
            Width           =   2775
         End
         Begin VB.TextBox txtcodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   2
            Left            =   2010
            MaxLength       =   10
            TabIndex        =   1
            Tag             =   "Cta.Contable|T|S|||sparam|ctaconta|||"
            Top             =   390
            Width           =   585
         End
         Begin VB.TextBox txtNombre 
            BackColor       =   &H80000018&
            Enabled         =   0   'False
            Height          =   285
            Index           =   4
            Left            =   2700
            Locked          =   -1  'True
            TabIndex        =   13
            Top             =   1140
            Width           =   2775
         End
         Begin VB.TextBox txtcodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   4
            Left            =   2010
            MaxLength       =   10
            TabIndex        =   3
            Tag             =   "Cta.Contable|T|S|||sparam|ctaconta|||"
            Top             =   1140
            Width           =   585
         End
         Begin VB.TextBox txtNombre 
            BackColor       =   &H80000018&
            Enabled         =   0   'False
            Height          =   285
            Index           =   5
            Left            =   2700
            Locked          =   -1  'True
            TabIndex        =   12
            Top             =   1530
            Width           =   2805
         End
         Begin VB.TextBox txtcodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   5
            Left            =   2010
            MaxLength       =   10
            TabIndex        =   4
            Tag             =   "Cta.Dif.negativas|T|S|||sparam|ctanegtat|||"
            Top             =   1530
            Width           =   585
         End
         Begin VB.TextBox txtcodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   3
            Left            =   2010
            MaxLength       =   10
            TabIndex        =   2
            Tag             =   "Código Postal|T|S|||clientes|codposta|||"
            Top             =   780
            Width           =   1050
         End
         Begin VB.Label Label1 
            Caption         =   "Número Diario "
            ForeColor       =   &H00972E0B&
            Height          =   195
            Index           =   0
            Left            =   270
            TabIndex        =   17
            Top             =   450
            Width           =   1350
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   2
            Left            =   1680
            ToolTipText     =   "Buscar Diario"
            Top             =   420
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Concepto al Haber"
            ForeColor       =   &H00972E0B&
            Height          =   195
            Index           =   25
            Left            =   270
            TabIndex        =   15
            Top             =   1530
            Width           =   1410
         End
         Begin VB.Label Label1 
            Caption         =   "Concepto al Debe"
            ForeColor       =   &H00972E0B&
            Height          =   195
            Index           =   24
            Left            =   270
            TabIndex        =   14
            Top             =   1170
            Width           =   1350
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   4
            Left            =   1680
            ToolTipText     =   "Buscar Concepto"
            Top             =   1170
            Width           =   240
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   5
            Left            =   1680
            ToolTipText     =   "Buscar Concepto"
            Top             =   1530
            Width           =   240
         End
         Begin VB.Image imgFec 
            Height          =   240
            Index           =   3
            Left            =   1680
            Picture         =   "frmContCieTurnoReg.frx":000C
            ToolTipText     =   "Buscar fecha"
            Top             =   810
            Width           =   240
         End
         Begin VB.Label Label4 
            Caption         =   "Fecha"
            ForeColor       =   &H00972E0B&
            Height          =   255
            Index           =   0
            Left            =   270
            TabIndex        =   11
            Top             =   810
            Width           =   495
         End
      End
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
         Height          =   1140
         Left            =   540
         TabIndex        =   8
         Top             =   210
         Width           =   5595
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
         Begin VB.Image imgFec 
            Height          =   240
            Index           =   0
            Left            =   1680
            Picture         =   "frmContCieTurnoReg.frx":0097
            ToolTipText     =   "Buscar fecha"
            Top             =   570
            Width           =   240
         End
         Begin VB.Label Label4 
            Caption         =   "Fecha"
            ForeColor       =   &H00972E0B&
            Height          =   255
            Index           =   16
            Left            =   300
            TabIndex        =   9
            Top             =   570
            Width           =   1425
         End
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   4905
         TabIndex        =   6
         Top             =   4860
         Width           =   975
      End
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   3720
         TabIndex        =   5
         Top             =   4860
         Width           =   975
      End
      Begin MSComctlLib.ProgressBar Pb1 
         Height          =   345
         Left            =   690
         TabIndex        =   18
         Top             =   3810
         Width           =   5370
         _ExtentX        =   9472
         _ExtentY        =   609
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Label lblProgres 
         Height          =   285
         Index           =   0
         Left            =   720
         TabIndex        =   20
         Top             =   4200
         Width           =   5265
      End
      Begin VB.Label lblProgres 
         Height          =   285
         Index           =   1
         Left            =   690
         TabIndex        =   19
         Top             =   4560
         Width           =   5295
      End
   End
End
Attribute VB_Name = "frmContCieTurnoReg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' +-+-+-+-+-+-+-+-+-+-+-
' +-+- Autor MONICA +-+-
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
Private WithEvents frmConce As frmConceConta 'conceptos de contabilidad
Attribute frmConce.VB_VarHelpID = -1
Private WithEvents frmTDia As frmDiaConta 'diarios de contabilidad
Attribute frmTDia.VB_VarHelpID = -1


'GENERALES PARA PASARLE A CRYSTAL REPORT
Private cadFormula As String 'Cadena con la FormulaSelection para Crystal Report
Private cadParam As String 'Cadena con los parametros para Crystal Report
Private numParam As Byte 'Numero de parametros que se pasan a Crystal Report
Private cadSelect As String 'Cadena para comprobar si hay datos antes de abrir Informe
Private cadTitulo As String 'Titulo para la ventana frmImprimir
Private cadNombreRPT As String 'Nombre del informe

Dim indCodigo As Integer 'indice para txtCodigo
 
'Se inicializan para cada Informe (tabla de BD a la que hace referencia
Dim Codigo As String 'Código para FormulaSelection de Crystal Report
Dim TipCod As String

Dim PrimeraVez As Boolean

Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, 0, cerrar
    If cerrar Then Unload Me
End Sub

Private Sub cmdAceptar_Click()
Dim sql As String
Dim i As Byte
Dim cadwhere As String

    If Not DatosOk Then Exit Sub
             
    sql = "SELECT count(*)" & _
          " FROM srecau, sforpa " & _
          "WHERE srecau.fechatur = " & DBSet(txtCodigo(0).Text, "F") & " and " & _
               " srecau.codforpa = sforpa.codforpa and " & _
               " srecau.intconta = 0 and " & _
               " sforpa.cuadresn = 1 and not sforpa.codmacta is null and mid(sforpa.codmacta,1,1) <> ' '"
             
    If RegistrosAListar(sql) = 0 Then
        MsgBox "No existen datos a contabilizar a esa fecha.", vbExclamation
        ' añadido, se han de marcar como contabilizados
        sql = "update srecau set intconta = 1 where srecau.fechatur = " & DBSet(txtCodigo(0).Text, "F")
        Conn.Execute sql
        Exit Sub
    End If
    
    cadwhere = " scaalb.fecalbar = " & DBSet(txtCodigo(0).Text, "F") & " and " & _
               " sforpa.contabilizasn = 1 "
               
    
    ContabilizarCierre (cadwhere)
     'Eliminar la tabla TMP
    BorrarTMPErrComprob

    DesBloqueoManual ("CIEREC") 'CIErre RECaudacion
    
    
    
eError:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "No se ha realizado el proceso de contabilización de cierre de turno. Llame a soporte."
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
     Me.imgBuscar(2).Picture = frmPpal.imgListImages16.ListImages(1).Picture
     Me.imgBuscar(4).Picture = frmPpal.imgListImages16.ListImages(1).Picture
     Me.imgBuscar(5).Picture = frmPpal.imgListImages16.ListImages(1).Picture
     txtCodigo(0).Text = Format(Now, "dd/mm/yyyy")
     
     txtCodigo(2).Text = Format(vParamAplic.NumDiario, "000")
     txtCodigo(4).Text = Format(vParamAplic.ConceptoDebe, "000")
     txtCodigo(5).Text = Format(vParamAplic.ConceptoHaber, "000")
     txtNombre(2).Text = DevuelveDesdeBDNew(cConta, "tiposdiario", "desdiari", "numdiari", txtCodigo(2).Text, "N")
     txtNombre(4).Text = PonerNombreConcepto(txtCodigo(4))
     txtNombre(5).Text = PonerNombreConcepto(txtCodigo(5))
     
    '###Descomentar
'    CommitConexion
         
    FrameCobrosVisible True, h, w
    Pb1.visible = False
    
    'Esto se consigue poneinedo el cancel en el opcion k corresponda
    Me.cmdCancel.Cancel = True
'   Me.Width = w + 70
'   Me.Height = h + 350
End Sub

Private Sub frmC_Selec(vFecha As Date)
 'Fecha
    txtCodigo(CByte(imgFec(0).Tag)).Text = Format(vFecha, "dd/MM/yyyy")
End Sub

Private Sub frmTDia_DatoSeleccionado(CadenaSeleccion As String)
'Form de Consulta de Clientes
    txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "00")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmConce_DatoSeleccionado(CadenaSeleccion As String)
'Form de Consulta de Clientes
    txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000")
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
    imgFec(0).Tag = Index 'independentment de les dates que tinga, sempre pose l'index en la 27
    If txtCodigo(Index).Text <> "" Then frmC.NovaData = txtCodigo(Index).Text

    frmC.Show vbModal
    Set frmC = Nothing
    PonerFoco txtCodigo(CByte(imgFec(0).Tag) + 1)
    ' ***************************
End Sub

Private Sub ListView1_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub imgBuscar_Click(Index As Integer)
   Select Case Index
        Case 2 ' TIPOS DE DIARIO
            AbrirFrmDiario (Index)
        
        Case 4, 5 'CONCEPTOS CONTABLES
            AbrirFrmConceptos (Index)
        
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
'14/02/2007
'    KEYpress KeyAscii
' ahora
    If KeyAscii = teclaBuscar Then
        Select Case Index
            Case 2: KEYBusqueda KeyAscii, 2 'diario
            Case 4: KEYBusqueda KeyAscii, 4 'concepto al debe
            Case 5: KEYBusqueda KeyAscii, 5 'concepto al haber
            Case 0: KEYFecha KeyAscii, 0 'fecha de turno
            Case 3: KEYFecha KeyAscii, 3 'fecha de contabilizacion
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
        Case 2 ' NUMERO DE DIARIO
            If txtCodigo(Index).Text <> "" Then
                txtNombre(Index).Text = ""
                txtNombre(Index).Text = DevuelveDesdeBDNew(cConta, "tiposdiario", "desdiari", "numdiari", txtCodigo(Index).Text, "N")
                If txtNombre(Index).Text = "" Then
                    MsgBox "Número de Diario no existe en la contabilidad. Reintroduzca.", vbExclamation
'                    PonerFoco txtcodigo(Index)
                End If
            End If
        
        Case 4, 5 'CONCEPTOS
            If txtCodigo(Index).Text <> "" Then txtNombre(Index).Text = PonerNombreConcepto(txtCodigo(Index))
            If txtNombre(Index).Text = "" Then
                MsgBox "Número de Concepto no existe en la contabilidad. Reintroduzca.", vbExclamation
'                PonerFoco txtcodigo(Index)
            End If

        Case 0, 3 'FECHAS
            If txtCodigo(Index).Text <> "" Then PonerFormatoFecha txtCodigo(Index)
            '26/03/2007 cuando cambien la fecha del cierre cambia la del asiento
            If Index = 0 Then 'And txtcodigo(3).Text = "" Then
                txtCodigo(3).Text = txtCodigo(0).Text
            End If
            
    End Select
End Sub

Private Sub FrameCobrosVisible(visible As Boolean, ByRef h As Integer, ByRef w As Integer)
    Me.FrameCobros.visible = visible
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

Private Sub AbrirFrmDiario(indice As Integer)
    indCodigo = indice
    Set frmTDia = New frmDiaConta
    frmTDia.DatosADevolverBusqueda = "0|1|"
    frmTDia.CodigoActual = txtCodigo(indCodigo)
    frmTDia.Show vbModal
    Set frmTDia = Nothing
End Sub

Private Sub AbrirFrmConceptos(indice As Integer)
    indCodigo = indice
    Set frmConce = New frmConceConta
    frmConce.DatosADevolverBusqueda = "0|1|"
    frmConce.CodigoActual = txtCodigo(indCodigo)
    frmConce.Show vbModal
    Set frmConce = Nothing
End Sub
 
Private Sub ContabilizarCierre(cadwhere As String)
'Contabiliza Facturas de Clientes o de Proveedores
Dim sql As String
Dim b As Boolean
Dim tmpErrores As Boolean 'Indica si se creo correctamente la tabla de errores
Dim CCoste As String
Dim cadTABLA As String

    sql = "CIEREC" 'contabilizar CIERRE DE RECAUDACION

    'Bloquear para que nadie mas pueda contabilizar
    DesBloqueoManual (sql)
    If Not BloqueoManual(sql, "1") Then
        MsgBox "No se pueden Contabilizar Cierre de Recaudación. Hay otro usuario contabilizándolo.", vbExclamation
        Screen.MousePointer = vbDefault
        Exit Sub
    End If

    
    'Visualizar la barra de Progreso
    Me.Pb1.visible = True
'    Me.Pb1.Top = 3350
    
    
    '==========================================================
    'REALIZAR COMPROBACIONES ANTES DE CONTABILIZAR FACTURAS
    '==========================================================
    
    Me.lblProgres(0).Caption = "Comprobaciones: "
    CargarProgres Me.Pb1, 100
        
    ' nuevo
    b = CrearTMPErrComprob()
    If Not b Then Exit Sub
    
    'comprobar que todas las CUENTAS de las distintos clientes que vamos a
    'contabilizar existen en la Conta: ssocio.codmacta IN (conta.cuentas.codmacta)
    '-----------------------------------------------------------------------------
    Me.MousePointer = vbHourglass

    Me.lblProgres(1).Caption = "Comprobando Cuentas Contables en contabilidad ..."
    cadTABLA = "ssocio"
    b = ComprobarCtaContable(cadTABLA, 1, cadwhere)
    IncrementarProgres Me.Pb1, 20
    Me.Refresh
    If Not b Then
        frmMensaje.OpcionMensaje = 2
        frmMensaje.Show vbModal
        Me.MousePointer = vbDefault
        Exit Sub
    End If
    
    Me.MousePointer = vbDefault
    
    'comprobar que todas las CUENTAS de las formas de pago que vamos a
    'contabilizar existen en la Conta: sforpa.ctaventa IN (conta.cuentas.codmacta)
    '-----------------------------------------------------------------------------
    Me.lblProgres(1).Caption = "Comprobando Cuentas Ctbles de Pago en contabilidad ..."
    b = ComprobarCtaContable(cadTABLA, 5)
    IncrementarProgres Me.Pb1, 20
    Me.Refresh
    If Not b Then
        frmMensaje.OpcionMensaje = 2
        frmMensaje.Show vbModal
        Exit Sub
    End If
    
'    'comprobar que todas las CUENTAS de diferencias positivas existen
'    'en la Conta: sparam.ctaposit IN (conta.cuentas.codmacta)
'    '-----------------------------------------------------------------------------
'    Me.lblProgres(1).Caption = "Comprobando Cuenta Ctble Diferencias Positivas en contabilidad ..."
'    b = ComprobarCtaContable(cadTABLA, 6)
'    IncrementarProgres Me.Pb1, 20
'    Me.Refresh
'    If Not b Then
'        frmMensaje.OpcionMensaje = 2
'        frmMensaje.Show vbModal
'        Exit Sub
'    End If
'    'comprobar que todas las CUENTAS de diferencias negativas existen
'    'en la Conta: sparam.ctanegtat IN (conta.cuentas.codmacta)
'    '-----------------------------------------------------------------------------
'    Me.lblProgres(1).Caption = "Comprobando Cuenta Ctble Diferencias Negativas en contabilidad ..."
'    b = ComprobarCtaContable(cadTABLA, 7)
'    IncrementarProgres Me.Pb1, 20
'    Me.Refresh
'    If Not b Then
'        frmMensaje.OpcionMensaje = 2
'        frmMensaje.Show vbModal
'        Exit Sub
'    End If
'
    
    '===========================================================================
    'CONTABILIZAR CIERRE
    '===========================================================================
    Me.lblProgres(0).Caption = "Contabilizar Cierre: "
    CargarProgres Me.Pb1, 10
    Me.lblProgres(1).Caption = "Insertando Asiento en Contabilidad..."
    
    
    cadwhere = "fechatur = " & DBSet(txtCodigo(0).Text, "F")
    b = PasarCierreAContab(cadwhere)
    
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
    
End Sub

Private Function DatosOk() As Boolean
Dim b As Boolean
Dim Orden1 As String
Dim Orden2 As String
Dim FFin As Date

   b = True

   If txtCodigo(0).Text = "" And b Then
        MsgBox "Introduzca la Fecha de recaudación a contabilizar.", vbExclamation
        b = False
        PonerFoco txtCodigo(0)
    End If
    
    ' comprobamos que han introducido los datos de la contabilidad
    ' +++NUMERO DE DIARIO+++
    If txtCodigo(2).Text = "" And b Then
        MsgBox "Introduzca Nº de diario a contabilizar.", vbExclamation
        b = False
        PonerFoco txtCodigo(2)
    End If
    
    ' +++FECHA DE ENTRADA+++
    If txtCodigo(3).Text = "" And b Then
        MsgBox "Introduzca la fecha de entrada del asiento.", vbExclamation
        b = False
        PonerFoco txtCodigo(3)
    Else
        ' comprobamos que la contabilizacion se encuentre en los ejercicios contables
         Orden1 = ""
         Orden1 = DevuelveDesdeBDNew(cConta, "parametros", "fechaini", "", "", "", "", "", "", "", "", "", "")
    
         Orden2 = ""
         Orden2 = DevuelveDesdeBDNew(cConta, "parametros", "fechafin", "", "", "", "", "", "", "", "", "", "")
         FFin = CDate(Orden2)
         If Not (CDate(Orden1) <= CDate(txtCodigo(3).Text) And CDate(txtCodigo(3).Text) < CDate(Day(FIni) & "/" & Month(FIni) & "/" & Year(FIni) + 2)) Then
            MsgBox "La Fecha de la contabilización no es del ejercicio actual ni del siguiente. Reintroduzca.", vbExclamation
            b = False
            PonerFoco txtCodigo(3)
         End If
    End If
    
    
    ' +++CONCEPTO AL DEBE+++
    If txtCodigo(4).Text = "" And b Then
        MsgBox "Introduzca el Concepto al Debe.", vbExclamation
        b = False
        PonerFoco txtCodigo(4)
    End If
    
    ' +++CONCEPTO AL HABER+++
    If txtCodigo(5).Text = "" And b Then
        MsgBox "Introduzca el Concepto al Haber.", vbExclamation
        b = False
        PonerFoco txtCodigo(5)
    End If
    
 
    DatosOk = b
End Function

Private Function PasarCierreAContab(cadwhere As String) As Boolean
Dim sql As String
Dim RS As ADODB.Recordset
Dim b As Boolean
Dim i As Integer
Dim NumLinea As Integer
Dim Mc As CContadorContab
Dim numdocum As String
Dim ampliacion As String
Dim ampliaciond As String
Dim ampliacionh As String
Dim ImporteD As Currency
Dim ImporteH As Currency
Dim Diferencia As Currency
Dim Obs As String
Dim cadMen As String
Dim cad As String
Dim CtaDifer As String


    On Error GoTo EPasarCie

    PasarCierreAContab = False
    
    'Total de lineas de asiento a Insertar en la contabilidad
    sql = "SELECT count(distinct scaalb.codforpa)" & _
          " FROM scaalb, sforpa " & _
          "WHERE scaalb.fecalbar = " & DBSet(txtCodigo(0).Text, "F") & " and " & _
               " scaalb.codforpa = sforpa.codforpa and " & _
               " sforpa.cuadresn = 1 and not sforpa.codmacta is null and mid(sforpa.codmacta,1,1) <> ' ' "

    NumLinea = TotalRegistros(sql)
    
    sql = "SELECT count(distinct ssocio.codmacta)" & _
          " FROM scaalb, sforpa, ssocio  " & _
          "WHERE scaalb.fecalbar = " & DBSet(txtCodigo(0).Text, "F") & " and " & _
               " scaalb.codforpa = sforpa.codforpa and " & _
               " scaalb.codsocio = ssocio.codsocio and " & _
               " sforpa.contabilizasn = 1 "
             
    NumLinea = NumLinea + TotalRegistros(sql)
    
    If NumLinea = 0 Then Exit Function
    
    If NumLinea > 0 Then
        NumLinea = NumLinea
        
        CargarProgres Me.Pb1, NumLinea
        
        ConnConta.BeginTrans
        Conn.BeginTrans
        
        Set Mc = New CContadorContab
        
        If Mc.ConseguirContador("0", (CDate(txtCodigo(3).Text) <= CDate(FFin)), True) = 0 Then
        
        Obs = "Cierre Turno de fecha " & Format(txtCodigo(0).Text, "dd/mm/yyyy")

    
        'Insertar en la conta Cabecera Asiento
        b = InsertarCabAsientoDia(txtCodigo(2).Text, Mc.Contador, txtCodigo(3).Text, Obs, cadMen)
        cadMen = "Insertando Cab. Asiento: " & cadMen
        
        If b Then
            sql = "SELECT scaalb.codforpa, sforpa.codmacta, sum(scaalb.importel)" & _
                  " FROM sforpa, scaalb " & _
                  " WHERE scaalb.fecalbar = " & DBSet(txtCodigo(0).Text, "F") & " and " & _
                        " scaalb.codforpa = sforpa.codforpa and " & _
                        " sforpa.cuadresn = 1 and not sforpa.codmacta is null and mid(sforpa.codmacta,1,1) <> ' '" & _
                  " GROUP BY scaalb.codforpa, sforpa.codmacta "
            
            Set RS = New ADODB.Recordset
            
            RS.Open sql, Conn, adOpenDynamic, adLockOptimistic, adCmdText
            
            i = 0
            ImporteD = 0
            ImporteH = 0
            
            numdocum = Format(CDate(txtCodigo(0).Text), "ddmmyy")
'            ampliacion = "Cierre Turno " & Format(txtcodigo(0).Text, "dd/mm/yyyy") & " T-" & Format(txtcodigo(1).Text, "0")
            ampliacion = "CTurno." & Format(txtCodigo(0).Text, "dd/mm/yy")
            ampliaciond = Trim(DevuelveDesdeBDNew(cConta, "conceptos", "nomconce", "codconce", txtCodigo(4).Text, "N")) & " " & ampliacion
            ampliacionh = Trim(DevuelveDesdeBDNew(cConta, "conceptos", "nomconce", "codconce", txtCodigo(5).Text, "N")) & " " & ampliacion
            
            
            If Not RS.EOF Then RS.MoveFirst
            While Not RS.EOF And b
                i = i + 1
                
                cad = DBSet(txtCodigo(2).Text, "N") & "," & DBSet(txtCodigo(3).Text, "F") & "," & DBSet(Mc.Contador, "N") & ","
                cad = cad & DBSet(i, "N") & "," & DBSet(RS!Codmacta, "T") & "," & DBSet(numdocum, "T") & ","
                
                ' COMPROBAMOS EL SIGNO DEL IMPORTE PQ NO PERMITIMOS INTRODUCIR APUNTES CON IMPORTES NEGATIVOS
                If RS.Fields(2).Value > 0 Then
                    ' importe al debe en positivo
                    cad = cad & DBSet(txtCodigo(4).Text, "N") & "," & DBSet(ampliaciond, "T") & "," & DBSet(RS.Fields(2).Value, "N") & ","
                    cad = cad & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & ",0"
                
                    ImporteD = ImporteD + CCur(RS.Fields(2).Value)
                Else
                    ' importe al haber en positivo, cambiamos el signo
                    cad = cad & DBSet(txtCodigo(5).Text, "N") & "," & DBSet(ampliacionh, "T") & "," & ValorNulo & ","
                    cad = cad & DBSet((RS.Fields(2).Value * -1), "N") & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & ",0"
                
                    ImporteH = ImporteH + (CCur(RS.Fields(2).Value) * (-1))
                End If
                
                cad = "(" & cad & ")"
                
                b = InsertarLinAsientoDia(cad, cadMen)
                cadMen = "Insertando Lin. Asiento: " & i
            
                IncrementarProgres Me.Pb1, 1
                Me.lblProgres(1).Caption = "Insertando línea de Asiento en Contabilidad...   (" & i & " de " & NumLinea & ")"
                Me.Refresh
            
                RS.MoveNext
            Wend
            RS.Close
            
            If b Then
                sql = "SELECT ssocio.codmacta, sum(importel) " & _
                      "FROM scaalb, ssocio, sforpa " & _
                      "WHERE scaalb.fecalbar = " & DBSet(txtCodigo(0).Text, "F") & " and " & _
                      " scaalb.codforpa = sforpa.codforpa and " & _
                      " scaalb.codsocio = ssocio.codsocio and " & _
                      " sforpa.contabilizasn = 1 " & _
                      " GROUP BY ssocio.codmacta"

                RS.Open sql, Conn, adOpenDynamic, adLockOptimistic, adCmdText
                
                
                If Not RS.EOF Then RS.MoveFirst
                
                
                While Not RS.EOF And b
                    i = i + 1
                
                    cad = DBSet(txtCodigo(2).Text, "N") & "," & DBSet(txtCodigo(3).Text, "F") & "," & DBSet(Mc.Contador, "N") & ","
                    cad = cad & DBSet(i, "N") & "," & DBSet(RS!Codmacta, "T") & "," & DBSet(numdocum, "T") & ","
                
                    ' COMPROBAMOS EL SIGNO DEL IMPORTE SI ES NEGATIVO LO PONEMOS EN EL DEBE CON SIGNO POSITIVO
                    If RS.Fields(1).Value > 0 Then
                        cad = cad & DBSet(txtCodigo(5).Text, "N") & "," & DBSet(ampliacionh, "T") & "," & ValorNulo & "," & DBSet(RS.Fields(1).Value, "N") & ","
                        cad = cad & ValorNulo & "," & ValorNulo & "," & ValorNulo & ",0"
                    
                        ImporteH = ImporteH + CCur(RS.Fields(1).Value)
                    
                    Else
                        cad = cad & DBSet(txtCodigo(4).Text, "N") & "," & DBSet(ampliaciond, "T") & "," & DBSet((RS.Fields(1).Value * -1), "N") & "," & ValorNulo & ","
                        cad = cad & ValorNulo & "," & ValorNulo & "," & ValorNulo & ",0"
                    
                        ImporteD = ImporteD + (CCur(RS.Fields(1).Value) * (-1))
                    
                    End If
                    cad = "(" & cad & ")"
                    
                    b = InsertarLinAsientoDia(cad, cadMen)
                    cadMen = "Insertando Lin. Asiento: " & i
                
                    IncrementarProgres Me.Pb1, 1
                    Me.lblProgres(1).Caption = "Insertando línea de Asiento en Contabilidad...   (" & i & " de " & NumLinea & ")"
                    Me.Refresh
                
                    RS.MoveNext
                Wend
                RS.Close
            
            
'                If b Then
'                    ' insertamos una linea al haber con la diferencia
'                    If ImporteD <> ImporteH Then
'                        i = i + 1
'
'                        If ImporteD > ImporteH Then
'                            Diferencia = ImporteD - ImporteH
'                            CtaDifer = vParamAplic.CtaPositiva
'                        Else
'                            Diferencia = ImporteH - ImporteD
'                            CtaDifer = vParamAplic.CtaNegativa
'                        End If
'
'                        cad = DBSet(txtcodigo(2).Text, "N") & "," & DBSet(txtcodigo(3).Text, "F") & "," & DBSet(Mc.Contador, "N") & ","
'                        cad = cad & DBSet(i, "N") & "," & DBSet(CtaDifer, "T") & "," & DBSet(numdocum, "T") & ","
'
'                        If ImporteD < ImporteH Then
'                            cad = cad & DBSet(txtcodigo(4).Text, "N") & "," & DBSet(ampliaciond, "T") & ","
'                            cad = cad & DBSet(Diferencia, "N") & "," & ValorNulo & ","
'                        Else
'                            cad = cad & DBSet(txtcodigo(5).Text, "N") & "," & DBSet(ampliacionh, "T") & ","
'                            cad = cad & ValorNulo & "," & DBSet(Diferencia, "N") & ","
'                        End If
'
'                        cad = cad & ValorNulo & "," & ValorNulo & "," & ValorNulo & ",0"
'
'                        cad = "(" & cad & ")"
'
'                        b = InsertarLinAsientoDia(cad, cadMen)
'                        cadMen = "Insertando Lin. Asiento: " & i
'
'                        IncrementarProgres Me.Pb1, 1
'                        Me.lblProgres(1).Caption = "Insertando línea de Asiento en Contabilidad...   (" & i & " de " & NumLinea & ")"
'                        Me.Refresh
'                    End If
'
'                End If

' de momento comentado para hacer pruebas
                If b Then
                    'Poner intconta=1 en arigasol.srecau
                    b = ActualizarRecaudacion(cadwhere, cadMen)
                    cadMen = "Actualizando Recaudación: " & cadMen
                End If
            
        End If
    End If
   End If
   End If
   
EPasarCie:
    If Err.Number <> 0 Then
        b = False
        MuestraError Err.Number, "Integrando Asiento a Contabilidad", Err.Description
    End If
    If b Then
        ConnConta.CommitTrans
        Conn.CommitTrans
        PasarCierreAContab = True
    Else
        ConnConta.RollbackTrans
        Conn.RollbackTrans
        PasarCierreAContab = False
    End If
End Function
