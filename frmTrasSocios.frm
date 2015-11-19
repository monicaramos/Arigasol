VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmTrasSocios 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Traspaso de Clientes"
   ClientHeight    =   3150
   ClientLeft      =   45
   ClientTop       =   2430
   ClientWidth     =   6735
   Icon            =   "frmTrasSocios.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3150
   ScaleWidth      =   6735
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
      Height          =   2955
      Left            =   90
      TabIndex        =   2
      Top             =   90
      Width           =   6555
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   660
         Top             =   2130
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         DefaultExt      =   "doc"
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   4965
         TabIndex        =   1
         Top             =   2070
         Width           =   975
      End
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   3750
         TabIndex        =   0
         Top             =   2070
         Width           =   975
      End
      Begin MSComctlLib.ProgressBar Pb1 
         Height          =   255
         Left            =   300
         TabIndex        =   3
         Top             =   1110
         Width           =   6030
         _ExtentX        =   10636
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Label Label1 
         Caption         =   "Proceso que importa los clientes que no existen en la aplicación."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   360
         TabIndex        =   6
         Top             =   390
         Width           =   5865
      End
      Begin VB.Label lblProgres 
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   5
         Top             =   1410
         Width           =   6075
      End
      Begin VB.Label lblProgres 
         Height          =   285
         Index           =   1
         Left            =   240
         TabIndex        =   4
         Top             =   1770
         Width           =   6075
      End
   End
End
Attribute VB_Name = "frmTrasSocios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' PROGRAMA DE TRASPASO DE SOCIOS ( SOLO PARA POBLA DEL DUC )
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
Dim cadTabla As String

Dim vContad As Long

Dim Socio As String
Dim Nombre As String
Dim NIF As String
Dim DIRECCION As String
Dim CPostal As String
Dim POBLACION As String
Dim Provincia As String
Dim IBAN As String
Dim CCC As String

Dim Banco As String
Dim Sucursal As String
Dim dc As String
Dim Cuenta As String


Dim PrimeraVez As Boolean

Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, 0, cerrar
    If cerrar Then Unload Me
End Sub


Private Sub cmdAceptar_Click()
Dim SQL As String
Dim I As Byte
Dim cadWhere As String
Dim b As Boolean
Dim NomFic As String
Dim Cadena As String
Dim cadena1 As String

On Error GoTo eError

    If Not DatosOk Then Exit Sub
    
    Me.CommonDialog1.Flags = cdlOFNExplorer + cdlOFNHideReadOnly + cdlOFNPathMustExist + cdlOFNFileMustExist


    Me.CommonDialog1.DefaultExt = "csv"
    'cadena = Format(CDate(txtcodigo(0).Text), FormatoFecha)
    CommonDialog1.FilterIndex = 1
    Me.CommonDialog1.FileName = "socios.csv"

    
    Me.CommonDialog1.CancelError = True
    
    Me.CommonDialog1.ShowOpen
    
    If Me.CommonDialog1.FileName <> "" Then
        InicializarVbles
            '========= PARAMETROS  =============================
        'Añadir el parametro de Empresa
        cadParam = cadParam & "|pEmpresa=""" & vEmpresa.nomEmpre & """|"
        numParam = numParam + 1


        If ProcesarFichero(Me.CommonDialog1.FileName) Then
            
            MsgBox "Proceso realizado correctamente.", vbExclamation
            Pb1.visible = False
            lblProgres(0).Caption = ""
            lblProgres(1).Caption = ""
            
            cadTabla = "tmpinformes"
            cadFormula = "{tmpinformes.codusu} = " & vSesion.Codigo
            
            SQL = "select count(*) from tmpinformes where codusu = " & vSesion.Codigo
            
            If TotalRegistros(SQL) <> 0 Then
                MsgBox "Han habido errores en el Traspaso de clientes. ", vbExclamation
                cadTitulo = "Errores en el Traspaso de clientes"
                cadNombreRPT = "rErroresTrasClientes.rpt"
                
                cadFormula = "{tmpinformes.codusu} = " & vSesion.Codigo
                
                LlamarImprimir
    
            End If
            cmdCancel_Click
        Else
            MsgBox "No se ha podido realizar el proceso.", vbExclamation
        End If
    
    Else
        MsgBox "No ha seleccionado ningún fichero", vbExclamation
        Exit Sub
    End If
             
             
eError:
    If Err.Number = 32755 Then Exit Sub ' le han dado a cancelar
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    If PrimeraVez Then
        PrimeraVez = False
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
    Pb1.visible = False
    
    'Esto se consigue poneinedo el cancel en el opcion k corresponda
    Me.cmdCancel.Cancel = True
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
Dim SQL As String
   
   b = True
   DatosOk = b
   
End Function



Private Function RecuperaFichero() As Boolean
Dim NF As Integer

    RecuperaFichero = False
    NF = FreeFile
    Open App.path For Input As #NF ' & "\BV" & Format(CDate(txtcodigo(0).Text), "ddmmyy") & "." & Format(txtcodigo(1).Text, "000") For Input As #NF
    Line Input #NF, cad
    Close #NF
    If cad <> "" Then RecuperaFichero = True
    
End Function

Private Sub InicializarTabla()
Dim SQL As String
    SQL = "delete from tmpinformes where codusu = " & vSesion.Codigo
    
    Conn.Execute SQL
End Sub

                
Private Function ProcesarFichero(nomFich As String) As Boolean
Dim NF As Long
Dim cad As String
Dim I As Integer
Dim longitud As Long
Dim RS As ADODB.Recordset
Dim RS1 As ADODB.Recordset
Dim Numreg As Long
Dim SQL As String
Dim Sql1 As String
Dim total As Long
Dim v_cant As Currency
Dim v_impo As Currency
Dim v_prec As Currency
Dim b As Boolean
Dim NomFic As String

    
    ProcesarFichero = False
    
    InicializarTabla
    
    NF = FreeFile
    
    Open nomFich For Input As #NF ' & "\BV" & Format(CDate(txtcodigo(0).Text), "ddmmyy") & "." & Format(txtcodigo(1).Text, "000") For Input As #NF
    
    Line Input #NF, cad ' saltamos la primera linea
    Line Input #NF, cad
    I = 1
    
    lblProgres(0).Caption = "Procesando Fichero: " & nomFich
    longitud = FileLen(nomFich)
    
    Pb1.visible = True
    Me.Pb1.Max = longitud
    Me.Refresh
    Me.Pb1.Value = 0
        
        
    b = True
    While Not EOF(NF) And b
        I = I + 1
        
        Me.Pb1.Value = Me.Pb1.Value + Len(cad)
        lblProgres(1).Caption = "Linea " & I
        Me.Refresh
        
        If cad <> ";;;;;;;;;" Then b = InsertarLinea(cad)
        
        Line Input #NF, cad
    Wend
    Close #NF
    
    If cad <> "" And b Then
        If cad <> ";;;;;;;;;" Then b = InsertarLinea(cad)
    End If
    
    ProcesarFichero = b
    
    Pb1.visible = False
    lblProgres(0).Caption = ""
    lblProgres(1).Caption = ""
    
End Function
                
                
                

            
Private Function InsertarLinea(cad As String) As Boolean
Dim SQL As String
Dim Codzona As String
Dim vSuperficie As Currency
Dim HayError As Boolean
Dim Mens As String
Dim Cadena As String

    On Error GoTo EInsertarLinea

    InsertarLinea = True
    
    
    CargarVariables cad
    
    'Comprobaciones para poder insertar
    
    'Comprobamos que el socio existe
    SQL = ""
    SQL = DevuelveDesdeBDNew(cPTours, "ssocio", "codsocio", "codsocio", Socio, "N")
    If SQL <> "" Then
        Mens = "Ya existe el cliente"
        Cadena = ""
        SQL = "insert into tmpinformes (codusu, importe1,  " & _
              " nombre2, nombre1) values (" & _
              vSesion.Codigo & "," & DBSet(Socio, "N") & ","
        SQL = SQL & DBSet(Cadena, "T") & "," & DBSet(Mens, "T") & ")"
        
        Conn.Execute SQL
        
        Exit Function
    End If
    
    
    If Socio = "" Then
        Mens = "No hay código de cliente"
        Cadena = ""
        SQL = "insert into tmpinformes (codusu, importe1,  " & _
              " nombre2, nombre1) values (" & _
              vSesion.Codigo & "," & DBSet(Socio, "N") & ","
        SQL = SQL & DBSet(Cadena, "T") & "," & DBSet(Mens, "T") & ")"
        
        Conn.Execute SQL
        
        Exit Function
    End If
    
    If Nombre = "" Then
        Mens = "Sin nombre de cliente"
        Cadena = ""
        SQL = "insert into tmpinformes (codusu, importe1,  " & _
              " nombre2, nombre1) values (" & _
              vSesion.Codigo & "," & DBSet(Socio, "N") & ","
        SQL = SQL & DBSet(Cadena, "T") & "," & DBSet(Mens, "T") & ")"
        
        Conn.Execute SQL
    
        Exit Function
    End If
    
    If NIF = "" Then
        Mens = "Sin nif de cliente"
        Cadena = ""
        SQL = "insert into tmpinformes (codusu, importe1,  " & _
              " nombre2, nombre1) values (" & _
              vSesion.Codigo & "," & DBSet(Socio, "N") & ","
        SQL = SQL & DBSet(Cadena, "T") & "," & DBSet(Mens, "T") & ")"
        
        Conn.Execute SQL
        Exit Function
    End If
    If DIRECCION = "" Then
        Mens = "Sin direccion"
        Cadena = ""
        SQL = "insert into tmpinformes (codusu, importe1,  " & _
              " nombre2, nombre1) values (" & _
              vSesion.Codigo & "," & DBSet(Socio, "N") & ","
        SQL = SQL & DBSet(Cadena, "T") & "," & DBSet(Mens, "T") & ")"
        
        Conn.Execute SQL
        Exit Function
    End If
    If CPostal = "" Then
        Mens = "Sin cpostal"
        Cadena = ""
        SQL = "insert into tmpinformes (codusu, importe1,  " & _
              " nombre2, nombre1) values (" & _
              vSesion.Codigo & "," & DBSet(Socio, "N") & ","
        SQL = SQL & DBSet(Cadena, "T") & "," & DBSet(Mens, "T") & ")"
        
        Conn.Execute SQL
        Exit Function
    End If
    If POBLACION = "" Then
        Mens = "Sin poblacion"
        Cadena = ""
        SQL = "insert into tmpinformes (codusu, importe1,  " & _
              " nombre2, nombre1) values (" & _
              vSesion.Codigo & "," & DBSet(Socio, "N") & ","
        SQL = SQL & DBSet(Cadena, "T") & "," & DBSet(Mens, "T") & ")"
        
        Conn.Execute SQL
        Exit Function
    End If
    If Provincia = "" Then
        Mens = "Sin provincia"
        Cadena = ""
        SQL = "insert into tmpinformes (codusu, importe1,  " & _
              " nombre2, nombre1) values (" & _
              vSesion.Codigo & "," & DBSet(Socio, "N") & ","
        SQL = SQL & DBSet(Cadena, "T") & "," & DBSet(Mens, "T") & ")"
        
        Conn.Execute SQL
        Exit Function
    End If
    If IBAN = "" And Len(CCC) <> 0 Then
        Mens = "Sin iban"
        Cadena = ""
        SQL = "insert into tmpinformes (codusu, importe1,  " & _
              " nombre2, nombre1) values (" & _
              vSesion.Codigo & "," & DBSet(Socio, "N") & ","
        SQL = SQL & DBSet(Cadena, "T") & "," & DBSet(Mens, "T") & ")"
        
        Conn.Execute SQL
        
    End If
    
    
    Banco = ""
    Sucursal = ""
    dc = ""
    Cuenta = ""
    If Len(CCC) <> 0 Then
        If Len(CCC) <> 20 Then
            Cadena = CCC
            Mens = "CCC errónea"
            SQL = "insert into tmpinformes (codusu, importe1,  " & _
                  "importe2, nombre2, nombre1) values (" & _
                  vSesion.Codigo & "," & DBSet(Socio, "N") & ","
            SQL = SQL & "0," & DBSet(Cadena, "T") & "," & DBSet(Mens, "T") & ")"
            
            Conn.Execute SQL
        Else
            Banco = Mid(CCC, 1, 4)
            Sucursal = Mid(CCC, 5, 4)
            dc = Mid(CCC, 9, 2)
            Cuenta = Mid(CCC, 11, 10)
        End If
    End If
    
    ' insertamos en la tabla de ssocios
    SQL = "insert into ssocio (codsocio,codcoope,nomsocio,domsocio,codposta,pobsocio,prosocio,"
    SQL = SQL & "nifsocio,fechaalt,codtarif,codsitua,codforpa,impfactu,iban,codbanco,codsucur,digcontr,cuentaba, "
    SQL = SQL & "codmacta) VALUES ("
    SQL = SQL & DBSet(Socio, "N") & ",1,"
    SQL = SQL & DBSet(Nombre, "T") & ","
    SQL = SQL & DBSet(DIRECCION, "T") & ","
    SQL = SQL & DBSet(CPostal, "T") & ","
    SQL = SQL & DBSet(POBLACION, "T") & ","
    SQL = SQL & DBSet(Provincia, "T") & ","
    SQL = SQL & DBSet(NIF, "T") & ","
    SQL = SQL & DBSet(Now, "F") & ","
    SQL = SQL & "1," ' tarifa
    SQL = SQL & "1," ' situacion
    SQL = SQL & "1," ' forma de pago
    SQL = SQL & "1," ' imprime factura
    SQL = SQL & DBSet(IBAN, "T") & ","
    SQL = SQL & DBSet(Banco, "T") & ","
    SQL = SQL & DBSet(Sucursal, "T") & ","
    SQL = SQL & DBSet(dc, "T") & ","
    SQL = SQL & DBSet(Cuenta, "T") & ","
    SQL = SQL & "'4333000')"
    
    
    Conn.Execute SQL
    
    ' insertamos en las tarjetas
    SQL = " insert into starje (codsocio,numlinea,numtarje,nomtarje,iban,codbanco,codsucur,digcontr,cuentaba,tiptarje) values ("
    SQL = SQL & DBSet(Socio, "N") & ",1," & DBSet(Socio, "N") & ","
    SQL = SQL & DBSet(Nombre, "T") & ","
    SQL = SQL & DBSet(IBAN, "T") & ","
    SQL = SQL & DBSet(Banco, "T") & ","
    SQL = SQL & DBSet(Sucursal, "T") & ","
    SQL = SQL & DBSet(dc, "T") & ","
    SQL = SQL & DBSet(Cuenta, "T") & ","
    SQL = SQL & "0)"
    
    Conn.Execute SQL
    
    
    Exit Function
    
EInsertarLinea:
    If Err.Number <> 0 Then
        InsertarLinea = False
        MsgBox "Error en Insertar Linea " & Err.Description, vbExclamation
    End If
End Function
            
            
            
Private Sub CargarVariables(cad As String)
    
    Socio = ""
    Nombre = ""
    NIF = ""
    DIRECCION = ""
    CPostal = ""
    POBLACION = ""
    Provincia = ""
    IBAN = ""
    CCC = ""

    Socio = RecuperaValorNew(cad, ";", 1)
    Nombre = RecuperaValorNew(cad, ";", 2)
    NIF = RecuperaValorNew(cad, ";", 3)
    DIRECCION = RecuperaValorNew(cad, ";", 4)
    CPostal = RecuperaValorNew(cad, ";", 5)
    POBLACION = RecuperaValorNew(cad, ";", 6)
    Provincia = RecuperaValorNew(cad, ";", 7)
    IBAN = RecuperaValorNew(cad, ";", 8)
    CCC = RecuperaValorNew(cad, ";", 9)
    
    

    
End Sub

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

