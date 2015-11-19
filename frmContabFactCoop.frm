VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmContabFactCoop 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Contabilizar Facturas de Cooperativas"
   ClientHeight    =   3690
   ClientLeft      =   45
   ClientTop       =   2430
   ClientWidth     =   6900
   Icon            =   "frmContabFactCoop.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3690
   ScaleWidth      =   6900
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
      Height          =   3465
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   6675
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   6
         Left            =   1605
         MaxLength       =   2
         TabIndex        =   2
         Top             =   1500
         Width           =   830
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   6
         Left            =   2505
         Locked          =   -1  'True
         TabIndex        =   9
         Text            =   "Text5"
         Top             =   1500
         Width           =   3645
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   3
         Left            =   1605
         MaxLength       =   10
         TabIndex        =   1
         Tag             =   "Código Postal|T|S|||clientes|codposta|||"
         Top             =   960
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
         Top             =   600
         Width           =   1050
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   5205
         TabIndex        =   4
         Top             =   2820
         Width           =   975
      End
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   4050
         TabIndex        =   3
         Top             =   2820
         Width           =   975
      End
      Begin MSComctlLib.ProgressBar Pb1 
         Height          =   255
         Left            =   360
         TabIndex        =   11
         Top             =   2190
         Width           =   5850
         _ExtentX        =   10319
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   420
         Top             =   2670
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Colectivo"
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
         Index           =   3
         Left            =   360
         TabIndex        =   10
         Top             =   1500
         Width           =   660
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   6
         Left            =   1305
         MouseIcon       =   "frmContabFactCoop.frx":000C
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar colectivo"
         Top             =   1500
         Width           =   240
      End
      Begin VB.Label Label4 
         Caption         =   "Fecha"
         ForeColor       =   &H00972E0B&
         Height          =   255
         Index           =   16
         Left            =   360
         TabIndex        =   8
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label Label4 
         Caption         =   "Desde"
         Height          =   195
         Index           =   15
         Left            =   720
         TabIndex        =   7
         Top             =   600
         Width           =   465
      End
      Begin VB.Label Label4 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   14
         Left            =   720
         TabIndex        =   6
         Top             =   960
         Width           =   420
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   2
         Left            =   1290
         Picture         =   "frmContabFactCoop.frx":015E
         ToolTipText     =   "Buscar fecha"
         Top             =   600
         Width           =   240
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   3
         Left            =   1290
         Picture         =   "frmContabFactCoop.frx":01E9
         ToolTipText     =   "Buscar fecha"
         Top             =   960
         Width           =   240
      End
   End
End
Attribute VB_Name = "frmContabFactCoop"
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
Private WithEvents frmCol As frmManCoope 'Colectivos
Attribute frmCol.VB_VarHelpID = -1
Private WithEvents frmTar As frmManTarje 'Tarjetas
Attribute frmTar.VB_VarHelpID = -1
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
Dim i As Byte
Dim sql As String
Dim nRegs As Long

    If Not DatosOk Then Exit Sub
    
    sql = "SELECT  count(*) "
    sql = sql & " from slhfacr, schfacr where 1 = 1 "
    If txtCodigo(2).Text <> "" Then sql = sql & " and slhfacr.fecfactu >= " & DBSet(txtCodigo(2).Text, "F")
    If txtCodigo(3).Text <> "" Then sql = sql & " and slhfacr.fecfactu <= " & DBSet(txtCodigo(3).Text, "F")
    sql = sql & " and schfacr.codcoope = " & DBSet(txtCodigo(6).Text, "N")
    sql = sql & " and slhfacr.letraser = schfacr.letraser and slhfacr.numfactu = schfacr.numfactu and slhfacr.fecfactu = schfacr.fecfactu "

    nRegs = TotalRegistros(sql)

    If nRegs <> 0 Then
        Pb1.visible = True
        Pb1.Max = nRegs
        Pb1.Value = 0
        
        If GeneraFichero Then
            If CopiarFichero Then
                MsgBox "Proceso realizado correctamente", vbExclamation
                cmdCancel_Click
            End If
        End If
    Else
        MsgBox "No hay datos entre esos límites.", vbExclamation
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

    'IMAGES para busqueda
     Me.imgBuscar(6).Picture = frmPpal.imgListImages16.ListImages(1).Picture

    '###Descomentar
'    CommitConexion
         
    FrameCobrosVisible True, h, w
    indFrame = 5
    tabla = "slhfac"
    
    Me.Pb1.visible = False
    
    'Esto se consigue poneinedo el cancel en el opcion k corresponda
    Me.cmdCancel.Cancel = True
End Sub

Private Sub frmC_Selec(vFecha As Date)
 'Fecha
    txtCodigo(CByte(imgFec(2).Tag)).Text = Format(vFecha, "dd/MM/yyyy")
End Sub

Private Sub frmFam_DatoSeleccionado(CadenaSeleccion As String)
'Form de Consulta de Familias
    txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "00")
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

Private Sub imgBuscar_Click(Index As Integer)
   Select Case Index
        Case 6 'COLECTIVO
            AbrirFrmColectivos (Index)
        
    End Select
    PonerFoco txtCodigo(indCodigo)
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
' ahora
    If KeyAscii = teclaBuscar Then
        Select Case Index
            Case 6: KEYBusqueda KeyAscii, 6 'colectivo
            Case 2: KEYFecha KeyAscii, 2 'fecha desde
            Case 3: KEYFecha KeyAscii, 3 'fecha hasta
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
        Case 2, 3 'FECHAS
            If txtCodigo(Index).Text <> "" Then PonerFormatoFecha txtCodigo(Index)
            
        Case 6 'COLECTIVO
            txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), "scoope", "nomcoope", "codcoope", "N")
            If txtCodigo(Index).Text <> "" Then txtCodigo(Index).Text = Format(txtCodigo(Index).Text, "00")
    End Select
End Sub

Private Sub FrameCobrosVisible(visible As Boolean, ByRef h As Integer, ByRef w As Integer)
    Me.FrameCobros.visible = visible
End Sub

Private Sub AbrirFrmColectivos(indice As Integer)
    indCodigo = indice
    Set frmCol = New frmManCoope
    frmCol.DatosADevolverBusqueda = "0|1|"
    frmCol.DeConsulta = True
    frmCol.CodigoActual = txtCodigo(indCodigo)
    frmCol.Show vbModal
    Set frmCol = Nothing
End Sub
 
Private Function GeneraFichero() As Boolean
Dim NFich1 As Integer
Dim NFich2 As Integer
Dim RS As ADODB.Recordset
Dim cad As String
Dim sql As String
Dim AntLetraser As String
Dim ActLetraser As String
Dim AntNumfactu As Long
Dim ActNumfactu As Long
Dim v_Hayreg As Integer
Dim AntTarjet As Long
Dim ActTarjet As Long
Dim AntFecfactu As Date
Dim ActFecfactu As Date

Dim vsocio As String

Dim NomSocio As String
Dim NomArtic As String
Dim b As Boolean
Dim Mens As String

    On Error GoTo EGen
    
    GeneraFichero = False

    NFich1 = FreeFile
    Open App.path & "\cabecera.txt" For Output As #NFich1
    NFich2 = FreeFile
    Open App.path & "\lineas.txt" For Output As #NFich2
    

    Set RS = New ADODB.Recordset
    
    'partimos de la tabla de historico de facturas
    sql = "SELECT slhfacr.letraser, slhfacr.numfactu, slhfacr.fecfactu, slhfacr.fecalbar, slhfacr.codartic, "
    sql = sql & " slhfacr.cantidad, slhfacr.preciove, slhfacr.implinea, ssocio.nrosocio, " 'slhfacr.numtarje, "
    sql = sql & " schfacr.codsocio "
    sql = sql & " from slhfacr, schfacr, ssocio where 1 = 1 "
    If txtCodigo(2).Text <> "" Then sql = sql & " and slhfacr.fecfactu >= " & DBSet(txtCodigo(2).Text, "F")
    If txtCodigo(3).Text <> "" Then sql = sql & " and slhfacr.fecfactu <= " & DBSet(txtCodigo(3).Text, "F")
    sql = sql & " and schfacr.codcoope = " & DBSet(txtCodigo(6).Text, "N")
    sql = sql & " and slhfacr.letraser = schfacr.letraser and slhfacr.numfactu = schfacr.numfactu and slhfacr.fecfactu = schfacr.fecfactu "
    sql = sql & " and ssocio.codsocio = schfacr.codsocio "
    sql = sql & " order by schfacr.codsocio, slhfacr.letraser, slhfacr.numfactu, slhfacr.fecfactu "

    RS.Open sql, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    
    AntLetraser = DBLet(RS!letraser, "T")
    ActLetraser = AntLetraser
    AntNumfactu = DBLet(RS!numfactu, "N")
    ActNumfactu = AntNumfactu
    AntFecfactu = DBLet(RS!fecfactu, "F")
    ActFecfactu = AntFecfactu
    AntTarjet = DBLet(RS!Nrosocio, "N") 'DBLet(RS!Numtarje)
    ActTarjet = AntTarjet
    
    b = True
    v_Hayreg = 0
    While Not RS.EOF And b
        v_Hayreg = 1
        
        ActLetraser = DBLet(RS!letraser)
        ActNumfactu = DBLet(RS!numfactu)
        ActFecfactu = DBLet(RS!fecfactu, "F")
        ActTarjet = DBLet(RS!Nrosocio, "N") 'DBLet(RS!Numtarje)
        
        Pb1.Value = Pb1.Value + 1
        
        vsocio = DBLet(RS!Nrosocio, "N") ' Mid(Format(DBLet(AntTarjet, "N"), "00000000"), 5, 4)
    
        If AntLetraser <> ActLetraser Or AntNumfactu <> ActNumfactu Or AntFecfactu <> ActFecfactu Then
            Mens = ""
            b = InsertarEnFichero1(NFich1, AntLetraser, AntNumfactu, CStr(AntTarjet), CStr(AntFecfactu), Mens)
            
            AntLetraser = ActLetraser
            AntNumfactu = ActNumfactu
            AntFecfactu = ActFecfactu
            AntTarjet = ActTarjet
        End If
    
        ' cargamos el fichero fichero2
        NomSocio = ""
        NomSocio = DevuelveDesdeBDNew(cPTours, "ssocio", "nomsocio", "codsocio", DBLet(RS!codsocio), "N")
        NomArtic = ""
        NomArtic = DevuelveDesdeBDNew(cPTours, "sartic", "nomartic", "codartic", DBLet(RS!codArtic), "N")
        
        '--- Laura
        'vsocio = Mid(Format(DBLet(RS!Numtarje, "N"), "00000000"), 5, 4)
        vsocio = DBLet(RS!Nrosocio, "N")
        '---
    
        cad = vsocio & "|"
        cad = cad & RellenaABlancos(NomSocio, True, 30) & "|"
'[Monica]02/07/2013: descomentamos la letra de serie, antes no la pasabamos
        cad = cad & DBLet(RS!letraser, "T")
        '[Monica]19/08/2013: si es llombai no separamos letra de nro de factura
        If CInt(txtCodigo(6).Text) = 1 Then
        
        Else
            cad = cad & "|"
        End If
'
        cad = cad & Format(DBLet(RS!numfactu, "N"), "0000000") & "|"
        cad = cad & Format(DBLet(RS!fecfactu, "F"), "dd/mm/yyyy") & "|"
        cad = cad & Format(DBLet(RS!fecAlbar, "F"), "dd/mm/yyyy") & "|"
        cad = cad & Format(DBLet(RS!codArtic, "N"), "000000") & "|"
        cad = cad & RellenaABlancos(NomArtic, True, 30) & "|"
        cad = cad & RellenaABlancos(Format(DBLet(RS!cantidad, "N"), "##,##0.00"), False, 9) & "|"
        cad = cad & RellenaABlancos(Format(DBLet(RS!preciove, "N"), "##0.00"), False, 6) & "|"
        cad = cad & RellenaABlancos(Format(DBLet(RS!ImpLinea, "N"), "##,##0.00"), False, 9) & "|"
    
        Print #NFich2, cad
            
        RS.MoveNext
    Wend
   
    If v_Hayreg = 1 And b Then
        ' metemos la última factura
        Mens = ""
        b = InsertarEnFichero1(NFich1, ActLetraser, ActNumfactu, vsocio, CStr(ActFecfactu), Mens)
    End If
    
EGen:
    Close (NFich1)
    Close (NFich2)
    Set RS = Nothing
    If Err.Number <> 0 Or Not b Then
        MuestraError Err.Number, Err.Description & vbCrLf & Mens
    Else
        GeneraFichero = True
    End If

End Function


Public Function CopiarFichero() As Boolean
Dim nomfich As String
Dim cadena As String
On Error GoTo ecopiarfichero

    CopiarFichero = False
    ' abrimos el commondialog para indicar donde guardarlo
'    Me.CommonDialog1.InitDir = App.path

    Me.CommonDialog1.DefaultExt = "txt"
    cadena = Format(txtCodigo(2).Text, FormatoFecha)
    CommonDialog1.Filter = "Archivos txt|txt|"
    CommonDialog1.FilterIndex = 1
    
    ' copiamos el primer fichero
    CommonDialog1.FileName = "cabecera.txt"
    
    Me.CommonDialog1.ShowSave
    If CommonDialog1.FileName <> "" Then
        FileCopy App.path & "\cabecera.txt", CommonDialog1.FileName
    End If
    
    'copiamos el segundo fichero
    CommonDialog1.FileName = "lineas.txt"
    
    Me.CommonDialog1.ShowSave
    If CommonDialog1.FileName <> "" Then
        FileCopy App.path & "\lineas.txt", CommonDialog1.FileName
    End If
    CopiarFichero = True


ecopiarfichero:
    If Err.Number <> 0 Then
        MuestraError Err.Number, Err.Description
    End If
    Err.Clear

End Function

Private Function RellenaABlancos(cadena As String, PorLaDerecha As Boolean, longitud As Integer) As String
Dim cad As String
    
    cad = Space(longitud)
    If PorLaDerecha Then
        cad = cadena & cad
        RellenaABlancos = Left(cad, longitud)
    Else
        cad = cad & cadena
        RellenaABlancos = Right(cad, longitud)
    End If
    
End Function

Private Function DatosOk() As Boolean

    DatosOk = True
    
    If txtCodigo(6).Text = "" Then
        MsgBox "Debe introducir un valor en el campo colectivo.", vbExclamation
        PonerFoco txtCodigo(6)
        DatosOk = False
        Exit Function
    End If
    
    
    
End Function

Private Function InsertarEnFichero1(NFich1 As Integer, letraser As String, numfactu As Long, vsocio As String, fecfactu As String, ByRef Mens As String) As Boolean
Dim sql As String
Dim RS As ADODB.Recordset
Dim vBase As Currency
Dim vIva As Currency
Dim cad As String


    On Error GoTo eInsertarEnFichero1

    InsertarEnFichero1 = False

    sql = "select schfacr.baseimp1, schfacr.baseimp2, schfacr.baseimp3, schfacr.impoiva1, "
    sql = sql & " schfacr.impoiva2, schfacr.impoiva3, schfacr.totalfac, schfacr.fecfactu "
    sql = sql & " from schfacr "
    sql = sql & " where letraser = " & DBSet(letraser, "T") & " and numfactu = " & DBSet(numfactu, "N")
    sql = sql & "  and fecfactu = " & DBSet(fecfactu, "F")
    
    Set RS = New ADODB.Recordset
    RS.Open sql, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    
    If Not RS.EOF Then
        vBase = DBLet(RS.Fields(0).Value, "N") + DBLet(RS.Fields(1).Value, "N") + DBLet(RS.Fields(2).Value, "N")
        vIva = DBLet(RS.Fields(3).Value, "N") + DBLet(RS.Fields(4).Value, "N") + DBLet(RS.Fields(5).Value, "N")
        
        ' cargamos el fichero fichero1
        cad = letraser ' & "|"
        '[Monica]19/08/2013: si es llombai no separamos letra de nro de factura
        If CInt(txtCodigo(6).Text) = 1 Then
        
        Else
            cad = cad & "|"
        End If
        
        
        cad = cad & Format(numfactu, "0000000") & "|"
        cad = cad & Format(DBLet(RS.Fields(7).Value, "F"), "dd/mm/yyyy") & "|"
        cad = cad & vsocio & "|"
        cad = cad & Format(vBase, "#######0.00") & "|"
        cad = cad & Format(vIva, "#######0.00") & "|"
        cad = cad & Format(DBLet(RS.Fields(6).Value, "N"), "#######0.00") & "|"
        Print #NFich1, cad
    End If
    InsertarEnFichero1 = True
eInsertarEnFichero1:
    If Err.Number <> 0 Then
        Mens = "Error en la Insercion en el fichero1 " & Err.Description
    End If
End Function

