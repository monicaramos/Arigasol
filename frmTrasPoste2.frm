VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmTrasPoste2 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Traspaso Datos Poste"
   ClientHeight    =   4890
   ClientLeft      =   45
   ClientTop       =   2430
   ClientWidth     =   6825
   Icon            =   "frmTrasPoste2.frx":0000
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
         Left            =   3720
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
      Begin VB.Label lblProgres 
         Caption         =   "Procesando"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   555
         Index           =   0
         Left            =   240
         TabIndex        =   5
         Top             =   1080
         Width           =   6195
      End
      Begin VB.Label lblProgres 
         Caption         =   "Fichero"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   285
         Index           =   1
         Left            =   240
         TabIndex        =   4
         Top             =   1710
         Width           =   6195
      End
   End
End
Attribute VB_Name = "frmTrasPoste2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' PROGRAMA DE TRASPASO DE POSTES PARA EL REGAIXO
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
Dim NomFicheros(7) As String
Dim nompath As String
Dim LongitudCadena As Long

Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, 0, cerrar
    If cerrar Then Unload Me
End Sub


Private Sub cmdAceptar_Click()
Dim sql As String
Dim i As Byte
Dim cadwhere As String
Dim b As Boolean
Dim nomfic As String
Dim cadena As String
Dim cadena1 As String
Dim RS As ADODB.Recordset
Dim Sql2 As String
Dim Mens As String

On Error GoTo eError


'    nompath = GetFolder("Selecciona directorio")
     If Me.CommonDialog1.FileName <> "" Then
'    If nompath <> "" Then
        InicializarVbles
        InicializarTabla
            '========= PARAMETROS  =============================
        'Añadir el parametro de Empresa
        cadParam = cadParam & "|pEmpresa=""" & vEmpresa.nomEmpre & """|"
        numParam = numParam + 1

'        nompath = CargarPath(Me.CommonDialog1.FileName)

        LongitudCadena = 100

        Sql2 = "select * from tipocred order by tipocred "
        Set RS = New ADODB.Recordset
        RS.Open Sql2, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
        b = True
        While Not RS.EOF And b
            b = ComprobarErrores(nompath & "\credit" & Format(DBLet(RS!tipocred), "0") & ".cre", DBLet(RS!tipocred))
            RS.MoveNext
        Wend
'        If b Then b = ComprobarErrores(nompath & "\credit1.cre", 1)
'        If b Then b = ComprobarErrores(nompath & "\credit3.cre", 3)
'        If b Then b = ComprobarErrores(nompath & "\credit5.cre", 5)
'        If b Then b = ComprobarErrores(nompath & "\credit6.cre", 6)
'        If b Then b = ComprobarErrores(nompath & "\credit8.cre", 8)
        Set RS = Nothing
        If Not b Then Exit Sub
        
        sql = "select count(*) from tmpinformes where codusu = " & vSesion.Codigo
        
        If TotalRegistros(sql) <> 0 Then
            MsgBox "Hay errores en el Traspaso de Postes. Debe corregirlos previamente.", vbExclamation
            cadTitulo = "Errores de Traspaso de Poste"
            cadNombreRPT = "rErroresTrasPoste2.rpt"
            LlamarImprimir
            Exit Sub
        Else
            Conn.BeginTrans
            Sql2 = "select * from tipocred order by tipocred "
            Set RS = New ADODB.Recordset
            RS.Open Sql2, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
           
            b = True
            While Not RS.EOF And b
                Mens = ""
                b = ProcesarFichero(nompath & "\credit" & Format(DBLet(RS!tipocred), "0") & ".cre", DBLet(RS!tipocred), Mens)
                RS.MoveNext
            Wend
            Set RS = Nothing
'
'
'            b = ProcesarFichero(nompath & "\credit0.cre", 0)
'            If b Then b = ProcesarFichero(nompath & "\credit1.cre", 1)
'            If b Then b = ProcesarFichero(nompath & "\credit3.cre", 3)
'            If b Then b = ProcesarFichero(nompath & "\credit5.cre", 5)
'            If b Then b = ProcesarFichero(nompath & "\credit6.cre", 6)
'            If b Then b = ProcesarFichero(nompath & "\credit8.cre", 8)

            If b Then
                Mens = ""
                b = ProcesarFicheroTurnos(nompath & "\turnos.dat", Mens)
            End If
        End If
    End If

eError:
    If Err.Number <> 0 Or Not b Then
        Conn.RollbackTrans
        MsgBox "No se ha podido realizar el proceso. " & vbCrLf & vbCrLf & Mens, vbExclamation
    Else
        Conn.CommitTrans
        MsgBox "Proceso realizado correctamente.", vbExclamation
        Pb1.visible = False
        lblProgres(0).Caption = ""
        lblProgres(1).Caption = ""
        cmdCancel_Click
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub Form_Activate()

On Error GoTo eError
    
    If PrimeraVez Then
        PrimeraVez = False
    End If
    Screen.MousePointer = vbDefault

    Me.CommonDialog1.Flags = cdlOFNExplorer + cdlOFNAllowMultiselect + cdlOFNHideReadOnly + cdlOFNPathMustExist + cdlOFNFileMustExist

    Me.CommonDialog1.InitDir = App.path & "\tpv\regaixo"
    Me.CommonDialog1.DefaultExt = "cre"
    CommonDialog1.Filter = "Archivos CRE|*.cre|"
    CommonDialog1.FilterIndex = 1
    Me.CommonDialog1.FileName = "credit0.cre"
    Me.CommonDialog1.CancelError = True
    Me.CommonDialog1.ShowOpen
    lblProgres(0).Caption = "Directorio Seleccionado : "
    nompath = CargarPath(Me.CommonDialog1.FileName)
    lblProgres(1).Caption = "    " & nompath
eError:
    If Err.Number = cdlCancel Then Unload Me
End Sub

Private Sub Form_Load()
Dim h As Integer, w As Integer
Dim List As Collection



    PrimeraVez = True
    Limpiar Me

    'IMAGES para busqueda
    '###Descomentar
'    CommitConexion
         
    FrameCobrosVisible True, h, w
    Pb1.visible = False
    
    'Esto se consigue poneinedo el cancel en el opcion k corresponda
    Me.cmdCancel.Cancel = True
'   Me.Width = w + 70
'   Me.Height = h + 350

End Sub

Private Sub Form_Unload(Cancel As Integer)
    DesBloqueoManual ("TRASPOST")
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


Private Function RecuperaFichero() As Boolean
Dim nf As Integer

    RecuperaFichero = False
    nf = FreeFile
    Open App.path For Input As #nf ' & "\BV" & Format(CDate(txtcodigo(0).Text), "ddmmyy") & "." & Format(txtcodigo(1).Text, "000") For Input As #NF
    Line Input #nf, cad
    Close #nf
    If cad <> "" Then RecuperaFichero = True
    
End Function


Private Function ProcesarFichero(nomfich As String, Opcion As Integer, ByRef Mens As String) As Boolean
Dim nf As Long
Dim cad As String
Dim i As Integer
Dim longitud As Long
Dim RS As ADODB.Recordset
Dim Rs1 As ADODB.Recordset
Dim NumReg As Long
Dim sql As String
Dim Codclave As String
Dim Codforpa As String

Dim v_fechac As String
Dim v_codsoc As String
Dim v_codpro As String
Dim v_ticket As String
Dim v_cantid As String
Dim v_precio As String
Dim v_import As String
Dim v_horaca As String
Dim c_import As Currency
Dim v_tarjet As String
Dim v_turnos As String

    On Error GoTo eProcesarFichero

    ProcesarFichero = False
    nf = FreeFile

    Open nomfich For Input As #nf

    i = 0

    lblProgres(0).Caption = "Procesando Fichero: " & nomfich
    longitud = FileLen(nomfich)

    If longitud < LongitudCadena Then
        Close #nf
        ProcesarFichero = True
        Exit Function
    End If

    Pb1.visible = True
    Me.Pb1.Max = longitud
    Me.Refresh
    Me.Pb1.Value = 0
    
    Do While Not EOF(nf)
        cad = Input(LongitudCadena, #nf)
    
        i = i + 1
        
'        If i = 54 Then
'            MsgBox "hola"
'        End If
        
        Me.Pb1.Value = Me.Pb1.Value + LongitudCadena 'Len(cad)
        lblProgres(1).Caption = "Linea " & i
        Me.Refresh
    
    
'        If Not (Mid(cad, 1, 2) = "  " Or Mid(cad, 100, 1) <> "L" Or Mid(cad, 75, 2) = "00") Then
        If Not (Mid(cad, 1, 2) = "  " Or Mid(cad, 75, 2) = "00") Then
            
            v_fechac = Mid(cad, 7, 2) & "/" & Mid(cad, 5, 2) & "/" & "20" & Mid(cad, 3, 2)
            v_codpro = 10000 + Mid(cad, 75, 2)
            v_cantid = Mid(cad, 58, 8)
            v_precio = Mid(cad, 48, 5)
            v_import = Mid(cad, 66, 7)
            v_horaca = Mid(cad, 9, 2) & Mid(cad, 11, 2) & "00"
            c_import = CCur(v_import) / 100
            v_ticket = LTrim(Mid(cad, 77, 6))
            v_turnos = Mid(cad, 83, 4)
            
            Select Case Opcion
                Case 0
                    v_codsoc = 900002
                    If v_ticket = "        " Or IsNull(v_ticket) Then v_ticket = "MANUAL"
                    If Not EsHoraOK(v_horaca) Then v_horaca = "090000"
                    
                    v_tarjet = DevuelveDesdeBDNew(cPTours, "starje", "numtarje", "codsocio", v_codsoc, "N")
                    'v_tarjet = "EFECTIVO"
                Case 1, 5, 6, 8
                    v_tarjet = Mid(cad, 21, 8)
                    v_codsoc = ""
                    v_codsoc = DevuelveDesdeBDNew(cPTours, "starje", "codsocio", "numtarje", v_tarjet, "N")
                    If Not EsHoraOK(v_horaca) Then v_horaca = "210000"
                Case 3
                    v_codsoc = 900000
                    If v_ticket = "        " Or IsNull(v_ticket) Then v_ticket = "MANUAL"
                    If CCur(v_horaca) = 0 Or Not EsHoraOK(v_horaca) Then v_horaca = "000100"
                    
                    v_tarjet = DevuelveDesdeBDNew(cPTours, "starje", "numtarje", "codsocio", v_codsoc, "N")
                    'v_tarjet = "BANCARIA"
            End Select
            
            Codclave = SugerirCodigoSiguienteStr("scaalb", "codclave")

            '02/03/2007
            'añadido el tema de la forma de pago que antes no se insertaba
            Codforpa = ""
            Codforpa = DevuelveDesdeBDNew(cPTours, "ssocio", "codforpa", "codsocio", v_codsoc, "N")

            sql = "insert into scaalb(codclave, codsocio, numalbar, fecalbar, numtarje, codartic, " & _
                   "cantidad, preciove, importel, horalbar, codturno, codforpa) values (" & _
                   DBSet(Codclave, "N") & "," & _
                   DBSet(v_codsoc, "N") & "," & DBSet(v_ticket, "N") & "," & DBSet(v_fechac, "F") & "," & _
                   DBSet(v_tarjet, "N") & "," & DBSet(v_codpro, "N") & "," & v_cantid & "," & _
                   v_precio & "," & DBSet(c_import, "N") & "," & DBSet(v_fechac & " " & v_horaca, "FH") & "," & _
                   DBSet(v_turnos, "N") & "," & DBSet(Codforpa, "N") & ")"
            Conn.Execute sql
            
            ' 15/03/2007 Monica:  actualizamos el precio de articulo
            sql = "update sartic set preventa = " & v_precio & _
                  " where codartic = " & DBSet(v_codpro, "N")
            Conn.Execute sql
            
            
            
        End If
        
    Loop
    
    If cad <> "" Then ProcesarFichero = True

    Pb1.visible = False
    lblProgres(0).Caption = ""
    lblProgres(1).Caption = ""
    
    Close #nf
eProcesarFichero:
    If Err.Number <> 0 Then
        Mens = "Error en el proceso de fichero credit" & Format(Opcion, "0") & "  " & Err.Description
    End If
    
End Function
                
Private Function ProcesarFicheroTurnos(nomfich As String, ByRef Mens As String) As Boolean
Dim nf As Long
Dim cad As String

Dim i As Integer
Dim longitud As Long
Dim RS As ADODB.Recordset
Dim Rs1 As ADODB.Recordset
Dim NumReg As Long
Dim sql As String

Dim v_horaca As String
Dim v_fechac As String
Dim v_numtan As String
Dim v_numman As String
Dim v_tipcre As String
Dim v_import As String
Dim v_cantid As String
Dim v_conini As String
Dim v_confin As String
Dim v_tipreg As Integer
Dim v_turnos As String
Dim c_conini As Currency
Dim c_confin As Currency
Dim c_import As Currency
Dim v_codart As Currency

Dim NumLinea As Long
Dim Existe As String
Dim cad2 As String



    On Error GoTo eProcesarFicheroTurnos

    ProcesarFicheroTurnos = False
    nf = FreeFile

    Open nomfich For Input As #nf
     
    LongitudCadena = 64
    
    i = 0
    
    lblProgres(0).Caption = "Procesando Fichero de turnos: " & nomfich
    longitud = FileLen(nomfich)
    
    If longitud < LongitudCadena Then
        Close #nf
        ProcesarFicheroTurnos = True
        Exit Function
    End If

    Pb1.visible = True
    Me.Pb1.Max = longitud
    Me.Refresh
    Me.Pb1.Value = 0
    
    NumLinea = 0
    
    Do While Not EOF(nf)
        cad = Input(LongitudCadena, #nf)
        
        i = i + 1
        
        Me.Pb1.Value = Me.Pb1.Value + LongitudCadena 'Len(cad)
        lblProgres(1).Caption = "Linea " & i
        Me.Refresh
    
        If Mid(cad, 19, 2) = "01" Or Mid(cad, 19, 2) = "03" Or Mid(cad, 19, 2) = "07" Then
            
            v_numtan = 0
            v_numman = 0
            v_tipcre = 0
            v_import = 0
            v_cantid = 0
            v_conini = 0
            v_confin = 0
        
            c_import = 0
            c_confin = 0
            c_conini = 0
            
            v_fechac = Mid(cad, 13, 2) & "/" & Mid(cad, 11, 2) & "/" & "20" & Mid(cad, 9, 2)

            v_horaca = Mid(cad, 15, 2) & Mid(cad, 17, 2) & "00"
            If Not EsHoraOK(v_horaca) Then v_horaca = "210000"

            v_turnos = Mid(cad, 3, 4)
            Select Case Mid(cad, 19, 2)
                Case "01"
                    v_tipreg = 2
                    v_tipcre = Mid(cad, 21, 20)
                    v_import = Mid(cad, 49, 8)
                    c_import = CCur(v_import / 100)
                    
                Case "03"
                    v_tipreg = 1
                    v_numtan = Mid(cad, 21, 20)
                    v_cantid = Mid(cad, 57, 8)
                
                Case "07"
                    v_tipreg = 0
                    v_numman = Mid(cad, 21, 20)
                    v_numtan = Mid(cad, 57, 8)
                    v_conini = Mid(cad, 41, 8)
                    v_confin = Mid(cad, 49, 8)
                    c_confin = CCur(v_confin) / 100
                    c_conini = CCur(v_conini) / 100
            End Select
            
            '20/03/2007 : cogemos el numero de linea que haya + 1 (pq pueden haber introducidas a mano)
            ' no se empieza de 0 como hacia antes
'            NumLinea = NumLinea + 1
            cad2 = "fechatur = " & DBSet(v_fechac, "F") & " and codturno = " & DBSet(v_turnos, "N")
            NumLinea = SugerirCodigoSiguienteStr("sturno", "numlinea", cad2)
            
            
            ' 02032007 no estaba añadido el codigo de articulo en el insert de sturno
            If v_numtan = 0 Then
                v_codart = 0
            Else
                v_codart = 10000 + CCur(v_numtan)
            End If
            
            'existe tipo de credito
            Existe = ""
            Existe = DevuelveDesdeBDNew(cPTours, "tipocred", "tipocred", "tipocred", v_tipcre, "N")
            If Existe <> "" Then
                sql = "insert into sturno (fechatur,  codturno,  numlinea, tiporegi,  numtanqu,  nummangu, " & _
                      "tipocred,  importel,  litrosve,  containi,  contafin, codartic) values (" & _
                      DBSet(v_fechac, "F") & "," & DBSet(v_turnos, "N") & "," & DBSet(NumLinea, "N") & "," & DBSet(v_tipreg, "N") & "," & _
                      DBSet(v_numtan, "N") & "," & DBSet(v_numman, "N") & "," & DBSet(v_tipcre, "N") & "," & _
                      DBSet(c_import, "N") & "," & v_cantid & "," & DBSet(c_conini, "N") & "," & _
                      DBSet(c_confin, "N") & "," & DBSet(v_codart, "N") & ")"
        
                Conn.Execute sql
            End If
        End If
    Loop
    
    If cad <> "" Then ProcesarFicheroTurnos = True

    Pb1.visible = False
    lblProgres(0).Caption = ""
    lblProgres(1).Caption = ""
    Close #nf
eProcesarFicheroTurnos:
    If Err.Number <> 0 Then
        Mens = "Error en el proceso del fichero de turnos " & vbCrLf & Err.Description
    End If
    
End Function
                
                
                
                
                
Private Function ComprobarErrores(nomfich As String, Opcion As Integer) As Boolean
Dim nf As Long
Dim cad As String
Dim i As Integer
Dim J As Integer
Dim longitud As Long
Dim RS As ADODB.Recordset
Dim NumReg As Long
Dim sql As String
Dim SeProcesaLinea As Boolean
Dim v_fechac As String
Dim v_codsoc As String
Dim v_codpro As String
Dim v_ticket As String
Dim v_cantid As String
Dim v_precio As String
Dim v_import As String
Dim v_horaca As String
Dim c_import As Currency
Dim v_tarjet As String
Dim Mens As String
Dim Caracter As String


    On Error GoTo eComprobarErrores
    
    ComprobarErrores = False

    If Dir(nomfich) = "" Then
        MsgBox "No existe el fichero " & nomfich, vbExclamation
        Exit Function
    End If
    
    nf = FreeFile
    
    Open nomfich For Input As #nf
    
    i = 0
    
    lblProgres(0).Caption = "Comprobando Fichero: " & nomfich
    longitud = FileLen(nomfich)
    
    If longitud < LongitudCadena Then
        Close #nf
        ComprobarErrores = True
        Exit Function
    End If
    
    Pb1.visible = True
    Me.Pb1.Max = longitud
    Me.Refresh
    Me.Pb1.Value = 0

    Do While Not EOF(nf)
        cad = Input(LongitudCadena, #nf)
        
        i = i + 1
        
        Me.Pb1.Value = Me.Pb1.Value + LongitudCadena 'Len(cad)
        lblProgres(1).Caption = "Linea " & i
        Me.Refresh
        
        SeProcesaLinea = False
        'cargamos los valores a comprobar
        Select Case Opcion
            Case 0
                ' si es una linea para procesar
'                If Not (Mid(cad, 1, 2) = "  " Or Mid(cad, 100, 1) <> "L" Or Mid(cad, 75, 2) = "00") Then
                
                If Not (Mid(cad, 1, 2) = "  " Or Mid(cad, 75, 2) = "00") Then
                    SeProcesaLinea = True
                    
                    v_fechac = Mid(cad, 7, 2) & Mid(cad, 5, 2) & Mid(cad, 3, 2)
                    v_codsoc = 900002
                    v_codpro = 10000 + Mid(cad, 75, 2)
                    v_ticket = LTrim(Mid(cad, 77, 6))
                    If v_ticket = "        " Or IsNull(v_ticket) Then v_ticket = "MANUAL"
                    v_cantid = Mid(cad, 58, 8)
                    v_precio = Mid(cad, 48, 5)
                    v_import = Mid(cad, 66, 7)
                    v_horaca = Mid(cad, 9, 2) & Mid(cad, 11, 2) & "00"
                    c_import = CCur(v_import) / 100
                    If Not EsHoraOK(v_horaca) Then v_horaca = "090000"
                    v_tarjet = "EFECTIVO"
        
                End If
            Case 1
'                If Not (Mid(cad, 1, 2) = "  " Or Mid(cad, 100, 1) <> "L" Or Mid(cad, 75, 2) = "00") Then
                If Not (Mid(cad, 1, 2) = "  " Or Mid(cad, 75, 2) = "00") Then
                    SeProcesaLinea = True
                    
                    v_fechac = Mid(cad, 7, 2) & Mid(cad, 5, 2) & "20" & Mid(cad, 3, 2)
                    v_tarjet = Mid(cad, 21, 8)
                    v_cantid = Mid(cad, 58, 8)
                    v_precio = Mid(cad, 48, 5)
                    v_import = Mid(cad, 66, 7)
                    c_import = CCur(v_import) / 100
                    v_codpro = 10000 + Mid(cad, 75, 2)
                    v_ticket = LTrim(Mid(cad, 77, 6))
                    v_horaca = Mid(cad, 9, 2) & Mid(cad, 11, 2) & "00"
                    If Not EsHoraOK(v_horaca) Then v_horaca = "210000"
                End If
            
            Case 3
'                If Not (Mid(cad, 1, 2) = "  " Or Mid(cad, 100, 1) <> "L" Or Mid(cad, 75, 2) = "00") Then
                If Not (Mid(cad, 1, 2) = "  " Or Mid(cad, 75, 2) = "00") Then
                    SeProcesaLinea = True
                    v_fechac = Mid(cad, 7, 2) & Mid(cad, 5, 2) & "20" & Mid(cad, 3, 2)
                    v_codsoc = 900000
                    v_cantid = Mid(cad, 58, 8)
                    v_precio = Mid(cad, 48, 5)
                    v_import = Mid(cad, 66, 7)
                    c_import = CCur(v_import) / 100
                    v_codpro = 10000 + Mid(cad, 75, 2)
                    v_ticket = LTrim(Mid(cad, 77, 6))
                    v_horaca = Mid(cad, 9, 2) & Mid(cad, 11, 2) & "00"
                    If v_ticket = "        " Or IsNull(v_ticket) Then v_ticket = "MANUAL"
                    If Not EsHoraOK(v_horaca) Then v_horaca = "000100"
                    
                    v_tarjet = "BANCARIA"
                End If
        
            Case 5
'                If Not (Mid(cad, 1, 2) = "  " Or Mid(cad, 100, 1) <> "L" Or Mid(cad, 75, 2) = "00") Then
                If Not (Mid(cad, 1, 2) = "  " Or Mid(cad, 75, 2) = "00") Then
                    SeProcesaLinea = True
                    v_fechac = Mid(cad, 7, 2) & Mid(cad, 5, 2) & "20" & Mid(cad, 3, 2)
                    v_tarjet = Mid(cad, 21, 8)
                    v_cantid = Mid(cad, 58, 8)
                    v_precio = Mid(cad, 48, 5)
                    v_import = Mid(cad, 66, 7)
                    c_import = CCur(v_import) / 100
                    v_codpro = 10000 + Mid(cad, 75, 2)
                    v_ticket = LTrim(Mid(cad, 77, 6))
                    v_horaca = Mid(cad, 9, 2) & Mid(cad, 11, 2) & "00"
                    If Not EsHoraOK(v_horaca) Then v_horaca = "210000"
                End If
            
            Case 6
'                If Not (Mid(cad, 1, 2) = "  " Or Mid(cad, 100, 1) <> "L" Or Mid(cad, 75, 2) = "00") Then
                If Not (Mid(cad, 1, 2) = "  " Or Mid(cad, 75, 2) = "00") Then
                    SeProcesaLinea = True
                    v_fechac = Mid(cad, 7, 2) & Mid(cad, 5, 2) & "20" & Mid(cad, 3, 2)
                    v_tarjet = Mid(cad, 21, 8)
                    v_cantid = Mid(cad, 58, 8)
                    v_precio = Mid(cad, 48, 5)
                    v_import = Mid(cad, 66, 7)
                    c_import = CCur(v_import) / 100
                    v_codpro = 10000 + Mid(cad, 75, 2)
                    v_ticket = LTrim(Mid(cad, 77, 6))
                    v_horaca = Mid(cad, 9, 2) & Mid(cad, 11, 2) & "00"
                    If Not EsHoraOK(v_horaca) Or IsNull(v_horaca) Then v_horaca = "210000"
                End If
            
            Case 8
'                If Not (Mid(cad, 1, 2) = "  " Or Mid(cad, 100, 1) <> "L" Or Mid(cad, 75, 2) = "00") Then
                If Not (Mid(cad, 1, 2) = "  " Or Mid(cad, 75, 2) = "00") Then
                    SeProcesaLinea = True
                    
                    v_fechac = Mid(cad, 7, 2) & Mid(cad, 5, 2) & "20" & Mid(cad, 3, 2)
                    v_tarjet = Mid(cad, 21, 8)
                    v_cantid = Mid(cad, 58, 8)
                    v_precio = Mid(cad, 48, 5)
                    v_import = Mid(cad, 66, 7)
                    c_import = CCur(v_import) / 100
                    v_codpro = 10000 + Mid(cad, 75, 2)
                    v_ticket = LTrim(Mid(cad, 77, 6))
                    v_horaca = Mid(cad, 9, 2) & Mid(cad, 11, 2) & "00"
                    If Not EsHoraOK(v_horaca) Or IsNull(v_horaca) Then v_horaca = "210000"
                End If
                
        End Select
        
        If SeProcesaLinea Then
            'Comprobamos fechas
            If Not EsFechaOK(v_fechac) Then
                Mens = "Fecha incorrecta"
                sql = "insert into tmpinformes (codusu, importe1, fecha1, campo1, campo2, nombre2, importe2, " & _
                      "importe3, importe4, importe5, nombre1) values (" & _
                      vSesion.Codigo & "," & DBSet(v_ticket, "T") & "," & DBSet(v_fechac, "F") & "," & DBSet(Format(v_horaca, "hh"), "N")
                Select Case Opcion
                    Case 0, 3
                      sql = sql & "," & DBSet(Format(v_horaca, "mm"), "N") & "," & DBSet(v_codsoc, "N") & "," & DBSet(v_codpro, "N") & "," & _
                      v_cantid & "," & v_precio & "," & DBSet(c_import, "N") & "," & DBSet(Mens, "T") & ")"
                    Case 1, 5, 6, 8
                      sql = sql & "," & DBSet(Format(v_horaca, "mm"), "N") & "," & DBSet(v_tarjet, "N") & "," & DBSet(v_codpro, "N") & "," & _
                      v_cantid & "," & v_precio & "," & DBSet(c_import, "N") & "," & DBSet(Mens, "T") & ")"
                End Select
                
                Conn.Execute sql
            End If
            
            
            'Comprobamos que el articulo existe en sartic
            sql = ""
            sql = DevuelveDesdeBDNew(cPTours, "sartic", "codartic", "codartic", v_codpro, "N")
            If sql = "" Then
                Mens = "No existe el artículo"
                sql = "insert into tmpinformes (codusu, importe1, fecha1, campo1, campo2, nombre2, importe2, " & _
                      "importe3, importe4, importe5, nombre1) values (" & _
                      vSesion.Codigo & "," & DBSet(v_ticket, "T") & "," & DBSet(v_fechac, "F") & "," & DBSet(Format(v_horaca, "hh"), "N")
                Select Case Opcion
                    Case 0, 3
                      sql = sql & "," & DBSet(Format(v_horaca, "mm"), "N") & "," & DBSet(v_codsoc, "N") & "," & DBSet(v_codpro, "N") & "," & _
                      v_cantid & "," & v_precio & "," & DBSet(c_import, "N") & "," & DBSet(Mens, "T") & ")"
                    Case 1, 5, 6, 8
                      sql = sql & "," & DBSet(Format(v_horaca, "mm"), "N") & "," & DBSet(v_tarjet, "N") & "," & DBSet(v_codpro, "N") & "," & _
                      v_cantid & "," & v_precio & "," & DBSet(c_import, "N") & "," & DBSet(Mens, "T") & ")"
                End Select
                      
                Conn.Execute sql
            End If
            
            
            'Comprobamos que el socio existe
            If Opcion = 0 Or Opcion = 3 Then
                sql = ""
                sql = DevuelveDesdeBDNew(cPTours, "ssocio", "codsocio", "codsocio", v_codsoc, "N")
                If sql = "" Then
                    Mens = "No existe el cliente"
                    sql = "insert into tmpinformes (codusu, importe1, fecha1, campo1, campo2, nombre2, importe2, " & _
                          "importe3, importe4, importe5, nombre1) values (" & _
                          vSesion.Codigo & "," & DBSet(v_ticket, "T") & "," & DBSet(v_fechac, "F") & "," & DBSet(Format(v_horaca, "hh"), "N")
                          sql = sql & "," & DBSet(Format(v_horaca, "mm"), "N") & "," & DBSet(v_codsoc, "N") & "," & DBSet(v_codpro, "N") & "," & _
                          v_cantid & "," & v_precio & "," & DBSet(c_import, "N") & "," & DBSet(Mens, "T") & ")"
                          
                    Conn.Execute sql
                End If
            End If
            
            'Comprobamos que exista el numero de tarjeta bonificada y normal
            If Opcion = 1 Or Opcion = 5 Or Opcion = 6 Or Opcion = 8 Then
                If Mid(v_tarjet, 4, 1) = "0" Then
                'tarjeta normal
                    sql = DevuelveDesdeBDNew(cPTours, "starje", "numtarje", "numtarje", v_tarjet, "N", "tiptarje", 0, "N")
                    If sql = "" Then
                        Mens = "No de tarjeta normal no existe"
                        sql = "insert into tmpinformes (codusu, importe1, fecha1, campo1, campo2, nombre2, importe2, " & _
                              "importe3, importe4, importe5, nombre1) values (" & _
                              vSesion.Codigo & "," & DBSet(v_ticket, "T") & "," & DBSet(v_fechac, "F") & "," & DBSet(Format(v_horaca, "hh"), "N")
                              sql = sql & "," & DBSet(Format(v_horaca, "mm"), "N") & "," & DBSet(v_tarjet, "T") & "," & DBSet(v_codpro, "N") & "," & _
                              v_cantid & "," & v_precio & "," & DBSet(c_import, "N") & "," & DBSet(Mens, "T") & ")"
                              
                        Conn.Execute sql
                    End If
                Else
                'tarjeta bonificada
                    sql = DevuelveDesdeBDNew(cPTours, "starje", "numtarje", "numtarje", v_tarjet, "N", , "tiptarje", 1, "N")
                    If sql = "" Then
                        Mens = "No de tarjeta bonificada no existe"
                        sql = "insert into tmpinformes (codusu, importe1, fecha1, campo1, campo2, nombre2, importe2, " & _
                              "importe3, importe4, importe5, nombre1) values (" & _
                              vSesion.Codigo & "," & DBSet(v_ticket, "N") & "," & DBSet(v_fechac, "F") & "," & DBSet(Format(v_horaca, "hh"), "N")
                              sql = sql & "," & DBSet(Format(v_horaca, "mm"), "N") & "," & DBSet(v_tarjet, "T") & "," & DBSet(v_codpro, "N") & "," & _
                              v_cantid & "," & v_precio & "," & DBSet(c_import, "N") & "," & DBSet(Mens, "T") & ")"
                              
                        Conn.Execute sql
                    End If
                
                End If
            End If
            
            'Comprobamos que no exista ya el numalbar, fecalbar en la scaalb
            sql = ""
            sql = DevuelveDesdeBDNew(cPTours, "scaalb", "numalbar", "numalbar", v_ticket, "N", , "fecalbar", v_fechac, "F")
            If sql <> "" Then
                Mens = "El Número de ticket ya ha sido traspasado"
                sql = "insert into tmpinformes (codusu, importe1, fecha1, campo1, campo2, nombre2, importe2, " & _
                      "importe3, importe4, importe5, nombre1) values (" & _
                      vSesion.Codigo & "," & DBSet(v_ticket, "N") & "," & DBSet(v_fechac, "F") & "," & DBSet(Format(v_horaca, "hh"), "N")
                      sql = sql & "," & DBSet(Format(v_horaca, "mm"), "N") & "," & DBSet(v_tarjet, "T") & "," & DBSet(v_codpro, "N") & "," & _
                      v_cantid & "," & v_precio & "," & DBSet(c_import, "N") & "," & DBSet(Mens, "T") & ")"
                      
                Conn.Execute sql
            End If
        End If ' de se procesa linea
'        Line Input #NF, cad
    Loop
    Close #nf
    
    If cad <> "" Then ComprobarErrores = True
    
    Pb1.visible = False
    lblProgres(0).Caption = ""
    lblProgres(1).Caption = ""

eComprobarErrores:
    If Err.Number <> 0 Then
        cad = "Se ha producido un error en el proceso de comprobación"
        MsgBox cad, vbExclamation
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
Dim sql As String
    sql = "delete from tmpinformes where codusu = " & vSesion.Codigo
    
    Conn.Execute sql
End Sub


'Private Sub CargarFicheros(nomfich As String)
'Dim i As Integer
'Dim J As Integer
'Dim Ini As String
'Dim lon As Integer
'    'C:\programas\Arigasol\envios credit1.cre credit0.cre
'    'C:\programas\Arigasol\envios\credit1.cre
'
'    Ini = InStr(1, nomfich, ".cre")
'
'    ' recorremos el inicio de cadena hasta credit1.cre
'
'    While Asc(Mid(nomfich, Ini, 1)) <> 0 And Mid(nomfich, Ini, 1) <> "\"
'        Ini = Ini - 1
'    Wend
'    lon = Len(nomfich)
'    nomfich = Mid(nomfich, Ini + 1, lon)
'    ' en la cadena queda: credit1.cre credit0.cre o solo credit1
'    lon = Len(nomfich)
'    i = 0
'    While lon > 0
'        J = InStr(1, nomfich, ".cre") - 1
'        NomFicheros(i) = Mid(nomfich, 1, J)
'
'        nomfich = Mid(nomfich, J + 6, lon)
'        i = i + 1
'        lon = Len(nomfich)
'    Wend
'End Sub
'
Private Function CargarPath(nomfich As String) As String
'las cadenas de entrada pueden ser las dos siguientes:
' C:\programas\Arigasol\envios credit1.cre credit0.cre
' C:\programas\Arigasol\envios\credit1.cre
Dim i As Integer
Dim J As Integer
Dim Ini As String
Dim Ini2 As String

Dim lon As Integer

    Ini = InStr(1, nomfich, ".cre")
    Ini2 = InStr(1, nomfich, ".CRE")
    
    If Ini2 > Ini Then Ini = Ini2
    
    ' recorremos el inicio de cadena hasta credit1.cre
    While Asc(Mid(nomfich, Ini, 1)) <> 0 And Mid(nomfich, Ini, 1) <> "\"
        Ini = Ini - 1
    Wend

    CargarPath = Mid(nomfich, 1, Ini - 1)
End Function




'    Select Case Opcion
'        Case 0
'            If Not (Mid(cad, 1, 2) = "  " Or Mid(cad, 100, 1) <> "L" Or Mid(cad, 75, 2) = "00") Then
'                SeProcesaLinea = True
'                v_fechac = Mid(cad, 7, 2) & Mid(cad, 5, 2) & Mid(cad, 3, 2)
'                v_codsoc = 900002
'                v_codpro = 10000 + Mid(cad, 75, 2)
'                v_ticket = LTrim(Mid(cad, 77, 6))
'                If v_ticket = "        " Or IsNull(v_ticket) Then v_ticket = "MANUAL"
'                v_cantid = Mid(cad, 58, 8)
'                v_precio = Mid(cad, 48, 5)
'                v_import = Mid(cad, 66, 7)
'                v_horaca = Mid(cad, 9, 2) & Mid(cad, 11, 2) & "00"
'                c_import = CCur(v_import) / 100
'                If Not EsHoraOK(v_horaca) Then v_horaca = "090000"
'                v_tarjet = "EFECTIVO"
'                v_turnos = Mid(cad, 83, 4)
'            End If
'        Case 1
'            If Not (Mid(cad, 1, 2) = "  " Or Mid(cad, 100, 1) <> "L" Or Mid(cad, 75, 2) = "00") Then
'                SeProcesaLinea = True
'                v_fechac = Mid(cad, 7, 2) & Mid(cad, 5, 2) & "20" & Mid(cad, 3, 2)
'                v_tarjet = Mid(cad, 21, 8)
'                v_cantid = Mid(cad, 58, 8)
'                v_precio = Mid(cad, 48, 5)
'                v_import = Mid(cad, 66, 7)
'                c_import = CCur(v_import) / 100
'                v_codpro = 10000 + Mid(cad, 75, 2)
'                v_ticket = LTrim(Mid(cad, 77, 6))
'                v_horaca = Mid(cad, 9, 2) & Mid(cad, 11, 2) & "00"
'                If Not EsHoraOK(v_horaca) Then v_horaca = "210000"
'                v_turnos = Mid(cad, 83, 4)
'            End If
'        Case 3
'            If Not (Mid(cad, 1, 2) = "  " Or Mid(cad, 100, 1) <> "L" Or Mid(cad, 75, 2) = "00") Then
'                SeProcesaLinea = True
'                v_fechac = Mid(cad, 7, 2) & Mid(cad, 5, 2) & "20" & Mid(cad, 3, 2)
'                v_codsoc = 900000
'                v_cantid = Mid(cad, 58, 8)
'                v_precio = Mid(cad, 48, 5)
'                v_import = Mid(cad, 66, 7)
'                c_import = CCur(v_import) / 100
'                v_codpro = 10000 + Mid(cad, 75, 2)
'                v_ticket = LTrim(Mid(cad, 77, 6))
'                v_horaca = Mid(cad, 9, 2) & Mid(cad, 11, 2) & "00"
'                v_turnos = Mid(cad, 83, 4)
'                If v_ticket = "        " Or IsNull(v_ticket) Then v_ticket = "MANUAL"
'                If CCur(v_horaca) = 0 Or Not EsHoraOK(v_horaca) Then v_horaca = "000100"
'
'                v_tarjet = "BANCARIA"
'            End If
'        Case 5
'            If Not (Mid(cad, 1, 2) = "  " Or Mid(cad, 100, 1) <> "L" Or Mid(cad, 75, 2) = "00") Then
'                SeProcesaLinea = True
'                v_fechac = Mid(cad, 7, 2) & Mid(cad, 5, 2) & "20" & Mid(cad, 3, 2)
'                v_tarjet = Mid(cad, 21, 8)
'                v_cantid = Mid(cad, 58, 8)
'                v_precio = Mid(cad, 48, 5)
'                v_import = Mid(cad, 66, 7)
'                c_import = CCur(v_import) / 100
'                v_codpro = 10000 + Mid(cad, 75, 2)
'                v_ticket = LTrim(Mid(cad, 77, 6))
'                v_horaca = Mid(cad, 9, 2) & Mid(cad, 11, 2) & "00"
'                If Not EsHoraOK(v_horaca) Then v_horaca = "210000"
'                v_turnos = Mid(cad, 83, 4)
'            End If
'        Case 6
'            If v_horaca Is Null Then Let v_horaca = "210000"
'
'            If Not (Mid(cad, 1, 2) = "  " Or Mid(cad, 100, 1) <> "L" Or Mid(cad, 75, 2) = "00") Then
'                SeProcesaLinea = True
'                v_fechac = Mid(cad, 7, 2) & Mid(cad, 5, 2) & "20" & Mid(cad, 3, 2)
'                v_tarjet = Mid(cad, 21, 8)
'                v_cantid = Mid(cad, 58, 8)
'                v_precio = Mid(cad, 48, 5)
'                v_import = Mid(cad, 66, 7)
'                c_import = CCur(v_import) / 100
'                v_codpro = 10000 + Mid(cad, 75, 2)
'                v_ticket = LTrim(Mid(cad, 77, 6))
'                v_horaca = Mid(cad, 9, 2) & Mid(cad, 11, 2) & "00"
'                If Not EsHoraOK(v_horaca) Or IsNull(v_horaca) Then v_horaca = "210000"
'                v_turnos = Mid(cad, 83, 4)
'            End If
'        Case 8
'            If Not (Mid(cad, 1, 2) = "  " Or Mid(cad, 100, 1) <> "L" Or Mid(cad, 75, 2) = "00") Then
'                SeProcesaLinea = True
'
'                v_fechac = Mid(cad, 7, 2) & Mid(cad, 5, 2) & "20" & Mid(cad, 3, 2)
'                v_tarjet = Mid(cad, 21, 8)
'                v_cantid = Mid(cad, 58, 8)
'                v_precio = Mid(cad, 48, 5)
'                v_import = Mid(cad, 66, 7)
'                c_import = CCur(v_import) / 100
'                v_codpro = 10000 + Mid(cad, 75, 2)
'                v_ticket = LTrim(Mid(cad, 77, 6))
'                v_horaca = Mid(cad, 9, 2) & Mid(cad, 11, 2) & "00"
'                If Not EsHoraOK(v_horaca) Or IsNull(v_horaca) Then v_horaca = "210000"
'                v_turnos = Mid(cad, 83, 4)
'            End If
'    End Select
'
