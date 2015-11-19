VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmTrasPosteCast 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Traspaso Datos Poste"
   ClientHeight    =   4890
   ClientLeft      =   45
   ClientTop       =   2430
   ClientWidth     =   6825
   Icon            =   "frmTrasPosteCast.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4890
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
      Height          =   4665
      Left            =   150
      TabIndex        =   2
      Top             =   120
      Width           =   6555
      Begin VB.TextBox text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   17
         Left            =   1575
         MaxLength       =   4
         TabIndex        =   7
         Tag             =   "Código F.Pago|N|N|0|999|ssocio|codforpa|000||"
         Top             =   1005
         Width           =   555
      End
      Begin VB.TextBox text2 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   17
         Left            =   2175
         TabIndex        =   6
         Top             =   1005
         Width           =   3765
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   570
         Top             =   3390
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         DefaultExt      =   "doc"
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   4905
         TabIndex        =   1
         Top             =   3780
         Width           =   975
      End
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   3690
         TabIndex        =   0
         Top             =   3780
         Width           =   975
      End
      Begin MSComctlLib.ProgressBar Pb1 
         Height          =   255
         Left            =   210
         TabIndex        =   3
         Top             =   2730
         Width           =   6030
         _ExtentX        =   10636
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   0
         Left            =   1260
         ToolTipText     =   "Buscar F.Pago"
         Top             =   1005
         Width           =   240
      End
      Begin VB.Label Label11 
         Caption         =   "Forma Pago"
         ForeColor       =   &H00972E0B&
         Height          =   255
         Left            =   270
         TabIndex        =   8
         Top             =   990
         Width           =   930
      End
      Begin VB.Label lblProgres 
         Height          =   375
         Index           =   0
         Left            =   180
         TabIndex        =   5
         Top             =   3120
         Width           =   6195
      End
      Begin VB.Label lblProgres 
         Height          =   285
         Index           =   1
         Left            =   180
         TabIndex        =   4
         Top             =   3480
         Width           =   6195
      End
   End
End
Attribute VB_Name = "frmTrasPosteCast"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' PROGRAMA DE TRASPASO DE POSTE PARA CASTELLDUC
' basado en frmTrasPoste ( de Alzira )
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
Private WithEvents frmFpa As frmManFpago 'F.Pago
Attribute frmFpa.VB_VarHelpID = -1


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
Dim cad As String
Dim cadTABLA As String

Dim vContad As Long

Dim PrimeraVez As Boolean

Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, 0, cerrar
    If cerrar Then Unload Me
End Sub

Private Sub cmdAceptar_Click()
Dim Sql As String
Dim I As Byte
Dim cadwhere As String
Dim b As Boolean
Dim NomFic As String
Dim cadena As String
Dim cadena1 As String

On Error GoTo eError


    If Not DatosOK Then Exit Sub
    
    Me.CommonDialog1.DefaultExt = "TXT"
    CommonDialog1.FilterIndex = 1
    Me.CommonDialog1.FileName = "*.txt"
    Me.CommonDialog1.ShowOpen
    
    If Me.CommonDialog1.FileName <> "" Then
        InicializarVbles
        InicializarTabla
            '========= PARAMETROS  =============================
        'Añadir el parametro de Empresa
        cadParam = cadParam & "|pEmpresa=""" & vEmpresa.nomEmpre & """|"
        numParam = numParam + 1

          
          If ProcesarFichero2(Me.CommonDialog1.FileName) Then
                cadTABLA = "tmpinformes"
                cadFormula = "{tmpinformes.codusu} = " & vSesion.Codigo
                
                Sql = "select count(*) from tmpinformes where codusu = " & vSesion.Codigo
                
                If TotalRegistros(Sql) <> 0 Then
                    MsgBox "Hay errores en el Traspaso de Postes. Debe corregirlos previamente.", vbExclamation
                    cadTitulo = "Errores de Traspaso de Poste"
                    cadNombreRPT = "rErroresTrasPoste3.rpt"
                    LlamarImprimir
                    Exit Sub
                Else
                    Conn.BeginTrans
                    b = ProcesarFichero(Me.CommonDialog1.FileName)
                End If
          End If
    Else
        MsgBox "No ha seleccionado ningún fichero", vbExclamation
        Exit Sub
    End If
             
             
eError:
    If Err.Number <> 0 Or Not b Then
        Conn.RollbackTrans
        MsgBox "No se ha podido realizar el proceso. LLame a Ariadna.", vbExclamation
    Else
        Conn.CommitTrans
        MsgBox "Proceso realizado correctamente.", vbExclamation
        Pb1.visible = False
        lblProgres(0).Caption = ""
        lblProgres(1).Caption = ""
'        BorrarArchivo Me.CommonDialog1.FileName
        cmdCancel_Click
    End If
    
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    If PrimeraVez Then
        PrimeraVez = False
        PonerFoco text1(17)
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
Dim h As Integer, w As Integer
Dim List As Collection

    PrimeraVez = True
    limpiar Me

    'IMAGES para busqueda
    '###Descomentar
'    CommitConexion
         
    FrameCobrosVisible True, h, w
    Pb1.visible = False
    Me.imgBuscar(0).Picture = frmPpal.imgListImages16.ListImages(1).Picture
    'Esto se consigue poneinedo el cancel en el opcion k corresponda
    Me.cmdCancel.Cancel = True
'   Me.Width = w + 70
'   Me.Height = h + 350
End Sub

Private Sub Form_Unload(Cancel As Integer)
    DesBloqueoManual ("TRASPOST")
End Sub


Private Sub imgBuscar_Click(Index As Integer)
    TerminaBloquear
     Select Case Index
        Case 0 'formas de pago
            Set frmFpa = New frmManFpago
            frmFpa.DatosADevolverBusqueda = "0|1|"
            frmFpa.CodigoActual = text1(17).Text
            frmFpa.Show vbModal
            Set frmFpa = Nothing
            PonerFoco text1(17)
     End Select
     
End Sub

Private Sub frmFpa_DatoSeleccionado(CadenaSeleccion As String)
'Mantenimiento F.Pago
    text1(17).Text = RecuperaValor(CadenaSeleccion, 1) 'codforpa
    FormateaCampo text1(17)
    text2(17).Text = RecuperaValor(CadenaSeleccion, 2) 'nomforpa
End Sub



Private Sub ListView1_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub Text1_LostFocus(Index As Integer)
Dim cadMen As String
Dim Nuevo As Boolean

    
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub
    
    ' ***************** configurar els LostFocus dels camps de la capçalera *****************
    Select Case Index
        Case 17 'Forma de Pago
            If PonerFormatoEntero(text1(Index)) Then
                text2(Index) = PonerNombreDeCod(text1(Index), "sforpa", "nomforpa", "codforpa", "N")
                If text2(Index).Text = "" Then
                    MsgBox "No existe la forma de pago. Reintroduzca.", vbExclamation
                    PonerFoco text1(Index)
                End If
            Else
                text2(Index).Text = ""
            End If
        
            
    End Select
        ' ***************************************************************************
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = teclaBuscar Then
        Select Case Index
            Case 17: KEYBusqueda KeyAscii, 3 'forma pago
        End Select
    Else
        KEYpress KeyAscii
    End If
End Sub

Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'Alvançar/Retrocedir els camps en les fleches de desplaçament del teclat.
    KEYdown KeyCode
End Sub

Private Sub KEYBusqueda(KeyAscii As Integer, indice As Integer)
    KeyAscii = 0
    imgBuscar_Click (indice)
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

 

Private Function DatosOK() As Boolean
Dim b As Boolean
Dim Sql As String
   b = True

   If text1(17).Text = "" And b Then
        MsgBox "El campo forma de pago debe tener un valor. Reintroduzca.", vbExclamation
        b = False
        PonerFoco text1(17)
    End If
 
    DatosOK = b
End Function



Private Function RecuperaFichero() As Boolean
Dim nf As Integer

    RecuperaFichero = False
    nf = FreeFile
    Open App.path For Input As #nf ' & "\BV" & Format(CDate(txtcodigo(0).Text), "ddmmyy") & "." & Format(txtcodigo(1).Text, "000") For Input As #NF
    Line Input #nf, cad
    Close #nf
    If cad <> "" Then RecuperaFichero = True
    
End Function


Private Function ProcesarFichero(nomFich As String) As Boolean
Dim nf As Long
Dim cad As String
Dim I As Integer
Dim longitud As Long
Dim Rs As ADODB.Recordset
Dim Rs1 As ADODB.Recordset
Dim NumReg As Long
Dim Sql As String
Dim SQL1 As String
Dim Total As Long
Dim v_cant As Currency
Dim v_impo As Currency
Dim v_prec As Currency
Dim b As Boolean
Dim NomFic As String

    ProcesarFichero = False
    nf = FreeFile
    
    Open nomFich For Input As #nf ' & "\BV" & Format(CDate(txtcodigo(0).Text), "ddmmyy") & "." & Format(txtcodigo(1).Text, "000") For Input As #NF
    
    Line Input #nf, cad
    I = 0
    
    lblProgres(0).Caption = "Procesando Fichero: " & nomFich
    longitud = FileLen(nomFich)
    
    Pb1.visible = True
    Me.Pb1.Max = longitud
    Me.Refresh
    Me.Pb1.Value = 0
        
    b = True
    While Not EOF(nf)
        I = I + 1
        
        Me.Pb1.Value = Me.Pb1.Value + Len(cad)
        lblProgres(1).Caption = "Linea " & I
        Me.Refresh
        
        
        b = InsertarLinea(cad)
        
        If b = False Then
            ProcesarFichero = False
            Exit Function
        End If
        
        Line Input #nf, cad
    Wend
    Close #nf
    
    If cad <> "" Then
        b = InsertarLinea(cad)

        If b = False Then
            ProcesarFichero = False
            Exit Function
        End If
    End If
    
    ProcesarFichero = b
    
    Pb1.visible = False
    lblProgres(0).Caption = ""
    lblProgres(1).Caption = ""

End Function
                
Private Function ProcesarFichero2(nomFich As String) As Boolean
Dim nf As Long
Dim cad As String
Dim I As Integer
Dim longitud As Long
Dim Rs As ADODB.Recordset
Dim Rs1 As ADODB.Recordset
Dim NumReg As Long
Dim Sql As String
Dim SQL1 As String
Dim Total As Long
Dim v_cant As Currency
Dim v_impo As Currency
Dim v_prec As Currency
Dim b As Boolean

    On Error GoTo eProcesarFichero2
    
    ProcesarFichero2 = False
    
    nf = FreeFile
    Open nomFich For Input As #nf ' & "\BV" & Format(CDate(txtcodigo(0).Text), "ddmmyy") & "." & Format(txtcodigo(1).Text, "000") For Input As #NF
    
    Line Input #nf, cad
    I = 0
    
    lblProgres(0).Caption = "Insertando en Tabla temporal: " & nomFich
    longitud = FileLen(nomFich)
    
    Pb1.visible = True
    Me.Pb1.Max = longitud
    Me.Refresh
    Me.Pb1.Value = 0
    ' PROCESO DEL FICHERO VENTAS.TXT

    b = True

    While Not EOF(nf) And b
        I = I + 1
        
        Me.Pb1.Value = Me.Pb1.Value + Len(cad)
        lblProgres(1).Caption = "Linea " & I
        Me.Refresh
        
        b = ComprobarRegistro(cad)
        
        Line Input #nf, cad
    Wend
    Close #nf
    
    If cad <> "" Then
        I = I + 1
        
        Me.Pb1.Value = Me.Pb1.Value + Len(cad)
        lblProgres(1).Caption = "Linea " & I
        Me.Refresh
        
        b = ComprobarRegistro(cad)
    End If
    
    Pb1.visible = False
    lblProgres(0).Caption = ""
    lblProgres(1).Caption = ""

    ProcesarFichero2 = b
    Exit Function

eProcesarFichero2:
    ProcesarFichero2 = False
End Function
                

Private Function ComprobarRegistro(cad As String) As Boolean
Dim Sql As String

Dim Base As String
Dim NombreBase As String
Dim Turno As String
Dim NumAlbaran As String
Dim NumFactura As String
Dim IdVendedor As String
Dim NombreVendedor As String
Dim FechaHora As String
Dim CodigoCliente As String
Dim NombreCliente As String
Dim Matricula As String
Dim CodigoProducto As String
Dim Surtidor As String
Dim Manguera As String
Dim PrecioLitro As String
Dim Cantidad As String
Dim Importe As String
Dim Descuento As String
Dim IdTipoPago As String
Dim DescrTipoPago As String
Dim CodigoTipoPago As String
Dim NifCliente As String
Dim IdProducto As String
Dim Tarjeta As String
Dim Tarje As String


Dim c_Cantidad As Currency
Dim c_Importe As Currency
Dim c_Importe1 As Currency
Dim c_Importe2 As Currency
Dim c_Precio As Currency
Dim c_Precio2 As Currency
Dim c_Descuento As Currency

Dim Fecha As String
Dim Hora As String

Dim Mens As String


Dim codsoc As String

    On Error GoTo eComprobarRegistro

    ComprobarRegistro = True

    Turno = Mid(cad, 18, 1)
    
    NumAlbaran = "TICKET"
    
    Fecha = Mid(cad, 7, 2) & "/" & Mid(cad, 5, 2) & "/" & Mid(cad, 1, 4)
    Hora = Mid(cad, 9, 8)
    CodigoCliente = Mid(cad, 61, 4)
    Tarjeta = Mid(cad, 57, 8)
    
    '[Monica] 07/04/2010: el codigo de cliente lo sacamos de la tabla starje
    CodigoCliente = ""
    CodigoCliente = DevuelveDesdeBDNew(cPTours, "starje", "codsocio", "numtarje", Tarjeta, "T")
    
    IdProducto = 1
    
    PrecioLitro = Mid(cad, 28, 9)
    Cantidad = Mid(cad, 19, 9)
    Importe = Mid(cad, 37, 8)
    
    If Trim(Importe) = "" Then
        Exit Function
    Else
        If CCur(Importe) = 0 Then Exit Function
    End If
    
    c_Cantidad = CCur(TransformaPuntosComas(Cantidad))
    c_Importe = CCur(TransformaPuntosComas(Importe))
    c_Precio = CCur(TransformaPuntosComas(PrecioLitro))
    
    
    'Comprobamos fechas
    If Not EsFechaOK(Fecha) Then
            Mens = "Fecha incorrecta"
            Sql = "insert into tmpinformes (codusu, fecha1, campo1, campo2, importe2, nombre2, importe3, " & _
                  "importe4, importe5, nombre1) values (" & _
                  vSesion.Codigo & "," & DBSet(Fecha, "F") & "," & DBSet(Mid(Hora, 1, 2), "N")
            Sql = Sql & "," & DBSet(Mid(Hora, 4, 2), "N") & "," & DBSet(CodigoCliente, "N") & "," & DBSet(Fecha, "T") & "," & _
                  DBSet(c_Cantidad, "N") & "," & DBSet(c_Precio, "N") & "," & DBSet(c_Importe, "N") & "," & DBSet(Mens, "T") & ")"
            
            Conn.Execute Sql
    End If
    
    
    'Comprobamos que el articulo existe en sartic
    Sql = ""
    Sql = DevuelveDesdeBDNew(cPTours, "sartic", "codartic", "codartic", IdProducto, "N")
    If Sql = "" Then
        Mens = "No existe el artículo"
        Dim IdProducto1 As Currency
        IdProducto1 = CCur(IdProducto)
        Sql = "insert into tmpinformes (codusu, fecha1, campo1, campo2, importe2, nombre2, " & _
              "importe3, importe4, importe5, nombre1) values (" & _
              vSesion.Codigo & "," & DBSet(Fecha, "F") & "," & DBSet(Mid(Hora, 1, 2), "N")
        Sql = Sql & "," & DBSet(Mid(Hora, 4, 2), "N") & "," & DBSet(CodigoCliente, "N") & "," & DBSet(IdProducto1, "T") & "," & _
              DBSet(c_Cantidad, "N") & "," & DBSet(c_Precio, "N") & "," & DBSet(c_Importe, "N") & "," & DBSet(Mens, "T") & ")"
              
        Conn.Execute Sql
    End If
    
    
    'Comprobamos que el socio existe
    If CodigoCliente <> "" Then
        Sql = ""
        Sql = DevuelveDesdeBDNew(cPTours, "ssocio", "codsocio", "codsocio", CodigoCliente, "N")
        If Sql = "" Then
            Mens = "No existe el cliente"
            Sql = "insert into tmpinformes (codusu, fecha1, campo1, campo2, importe2, nombre2, importe3, " & _
                  "importe4, importe5, nombre1) values (" & _
                  vSesion.Codigo & "," & DBSet(Fecha, "F") & "," & DBSet(Mid(Hora, 1, 2), "N")
            Sql = Sql & "," & DBSet(Mid(Hora, 4, 2), "N") & "," & DBSet(CodigoCliente, "N") & "," & DBSet(CodigoCliente, "T") & "," & _
                    DBSet(c_Cantidad, "N") & "," & DBSet(c_Precio, "N") & "," & DBSet(c_Importe, "N") & "," & DBSet(Mens, "T") & ")"

            Conn.Execute Sql
        End If
    Else
    '[Monica] 07/04/2010: si no tiene socio asociado a la tarjeta
        Mens = "No existe el cliente"
        Sql = "insert into tmpinformes (codusu, fecha1, campo1, campo2, importe2, nombre2, importe3, " & _
              "importe4, importe5, nombre1) values (" & _
              vSesion.Codigo & "," & DBSet(Fecha, "F") & "," & DBSet(Mid(Hora, 1, 2), "N")
        Sql = Sql & "," & DBSet(Mid(Hora, 4, 2), "N") & "," & DBSet(CodigoCliente, "N") & "," & DBSet(Tarjeta, "T") & "," & _
                DBSet(c_Cantidad, "N") & "," & DBSet(c_Precio, "N") & "," & DBSet(c_Importe, "N") & "," & DBSet(Mens, "T") & ")"

        Conn.Execute Sql
    End If
    
    'Comprobamos que la tarjeta existe
    If Tarjeta <> "" Then
        Sql = ""
        Sql = DevuelveDesdeBDNew(cPTours, "starje", "codsocio", "numtarje", Tarjeta, "N")
        If Sql = "" Then
            Mens = "No existe la tarjeta"
            Sql = "insert into tmpinformes (codusu, fecha1, campo1, campo2, importe2, nombre2, importe3, " & _
                  "importe4, importe5, nombre1) values (" & _
                  vSesion.Codigo & "," & DBSet(Fecha, "F") & "," & DBSet(Mid(Hora, 1, 2), "N")
            Sql = Sql & "," & DBSet(Mid(Hora, 4, 2), "N") & "," & DBSet(CodigoCliente, "N") & "," & DBSet(Tarjeta, "T") & "," & _
                    DBSet(c_Cantidad, "N") & "," & DBSet(c_Precio, "N") & "," & DBSet(c_Importe, "N") & "," & DBSet(Mens, "T") & ")"

            Conn.Execute Sql
        End If
    End If
    
    
    
    'Comprobamos que la forma de pago existe
    If text1(17).Text <> "" Then
        Sql = ""
        Sql = DevuelveDesdeBDNew(cPTours, "sforpa", "codforpa", "forpaalvic", text1(17).Text, "N")
        If Sql = "" Then
            Mens = "No existe la forma de pago"
            Sql = "insert into tmpinformes (codusu, fecha1, campo1, campo2, importe2, nombre2, " & _
                  "importe3, importe4, importe5, nombre1) values (" & _
                  vSesion.Codigo & "," & DBSet(Fecha, "F") & "," & DBSet(Mid(Hora, 1, 2), "N")
            Sql = Sql & "," & DBSet(Mid(Hora, 4, 2), "N") & "," & DBSet(CodigoCliente, "N") & "," & DBSet(text1(17).Text, "T") & "," & _
                    DBSet(c_Cantidad, "N") & "," & DBSet(c_Precio, "N") & "," & DBSet(c_Importe, "N") & "," & DBSet(Mens, "T") & ")"
            
            Conn.Execute Sql
        End If
    End If
    
    
    'Comprobamos que la entrada no ha sido introducida
    Sql = "select count(*) from scaalb where fecalbar = " & DBSet(Fecha, "F")
    Sql = Sql & " and horalbar = " & DBSet(Fecha & " " & Hora, "FH")
    
    If TotalRegistros(Sql) > 0 Then
        Mens = "Existe la entrada"
        Sql = "insert into tmpinformes (codusu, fecha1, campo1, campo2, importe2, nombre2, " & _
              "importe3, importe4, importe5, nombre1) values (" & _
              vSesion.Codigo & "," & DBSet(Fecha, "F") & "," & DBSet(Mid(Hora, 1, 2), "N")
        Sql = Sql & "," & DBSet(Mid(Hora, 4, 2), "N") & "," & DBSet(CodigoCliente, "N") & "," & DBSet(NumAlbaran, "T") & "," & _
                DBSet(c_Cantidad, "N") & "," & DBSet(c_Precio, "N") & "," & DBSet(c_Importe, "N") & "," & DBSet(Mens, "T") & ")"
        
        Conn.Execute Sql
    End If
    
    
    
eComprobarRegistro:
    If Err.Number <> 0 Then
        ComprobarRegistro = False
    End If
End Function
            
            
            
Private Function InsertarLinea(cad As String) As Boolean
Dim NumLin As String
Dim codpro As String
Dim articulo As String
Dim Familia As String
Dim precio As String
Dim ImpDes As String
Dim CodIVA As String
Dim b As Boolean
Dim Codclave As String
Dim Sql As String

Dim Import As Currency

Dim Base As String
Dim NombreBase As String
Dim Turno As String
Dim NumAlbaran As String
Dim NumFactura As String
Dim IdVendedor As String
Dim NombreVendedor As String
Dim FechaHora As String
Dim Fecha As String
Dim Hora As String
Dim CodigoCliente As String
Dim NombreCliente As String
Dim Matricula As String
Dim Tarjeta As String
Dim CodigoProducto As String
Dim Surtidor As String
Dim Manguera As String
Dim PrecioLitro As String
Dim Cantidad As String
Dim Importe As String
Dim Descuento As String
Dim IdTipoPago As String
Dim DescrTipoPago As String
Dim CodigoTipoPago As String
Dim NifCliente As String
Dim IdProducto As String

Dim c_Cantidad As Currency
Dim c_Importe As Currency
Dim c_Importe1 As Currency
Dim c_Importe2 As Currency
Dim c_Precio As Currency
Dim c_Precio2 As Currency
Dim c_Descuento As Currency
Dim IdProductoDes As String

Dim Tarje As String


Dim Mens As String
Dim NumLinea As Long

Dim codsoc As String
Dim Forpa As String

    On Error GoTo eInsertarLinea

    InsertarLinea = True
    

    Turno = Mid(cad, 18, 1)
    
    NumAlbaran = "TICKET"
    
    Fecha = Mid(cad, 7, 2) & "/" & Mid(cad, 5, 2) & "/" & Mid(cad, 1, 4)
    Hora = Mid(cad, 9, 8)
    FechaHora = Format(Fecha, "yyyy-mm-dd") & " " & Hora
    
    CodigoCliente = Mid(cad, 61, 4)
    Tarjeta = Mid(cad, 57, 8)
    
    '[Monica] 07/04/2010: el codigo de cliente lo sacamos de la tabla starje
    CodigoCliente = ""
    CodigoCliente = DevuelveDesdeBDNew(cPTours, "starje", "codsocio", "numtarje", Tarjeta, "T")
    
    IdProducto = "1"
    
    PrecioLitro = Mid(cad, 28, 9)
    Cantidad = Mid(cad, 19, 9)
    Importe = Mid(cad, 37, 8)
    
    If Trim(Importe) = "" Then
        Exit Function
    Else
        If CCur(Importe) = 0 Then Exit Function
    End If
    
    c_Cantidad = CCur(TransformaPuntosComas(Cantidad))
    c_Importe = CCur(TransformaPuntosComas(Importe))
    c_Precio = CCur(TransformaPuntosComas(PrecioLitro))
    
'    If Trim(Importe) = "" Then
'        Exit Function
'    Else
'        If CCur(Importe) = 0 Then Exit Function
'    End If
    

    
'    '### [Monica] 17/09/2007
'    'no insertamos aquellas lineas de albaran de importe = 0
'    Importe = DBSet(c_Importe, "N")
'    If Import = 0 Then
'        InsertarLineaAlz = True
'        Exit Function
'    End If
'    'hasta aqui
    
    ' insertamos en la tabla de albaranes
    Codclave = SugerirCodigoSiguienteStr("scaalb", "codclave")
    Forpa = text1(17).Text
    
    Sql = "INSERT INTO scaalb (codclave, codsocio, numtarje, numalbar, fecalbar, horalbar, " & _
          "codturno, codartic, cantidad, preciove, importel, codforpa, matricul, codtraba, " & _
          "numfactu, numlinea) VALUES (" & DBSet(Codclave, "T") & "," & DBSet(CodigoCliente, "N") & "," & DBSet(Tarjeta, "N") & "," & _
           DBSet(NumAlbaran, "T") & "," & DBSet(Fecha, "F") & "," & DBSet(Fecha & " " & Hora, "FH") & "," & DBSet(Turno, "N") & "," & _
           DBSet(IdProducto, "N") & "," & DBSet(c_Cantidad, "N") & "," & DBSet(c_Precio, "N") & "," & _
           DBSet(c_Importe, "N") & "," & DBSet(text1(17).Text, "N") & "," & ValorNulo & ",1,"
    Sql = Sql & "0,0)"
    
    Conn.Execute Sql
        
 
    
    
eInsertarLinea:
    If Err.Number <> 0 Then
        InsertarLinea = False
        MsgBox "Error en Insertar Linea " & Err.Description, vbExclamation
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

Private Sub InicializarTabla()
Dim Sql As String
    Sql = "delete from tmpinformes where codusu = " & vSesion.Codigo
    
    Conn.Execute Sql
End Sub


