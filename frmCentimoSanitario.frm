VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmCentimoSanitario 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Centimo Sanitario"
   ClientHeight    =   6315
   ClientLeft      =   45
   ClientTop       =   2430
   ClientWidth     =   7185
   Icon            =   "frmCentimoSanitario.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6315
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
      Height          =   6015
      Left            =   120
      TabIndex        =   10
      Top             =   120
      Width           =   6915
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   6
         Left            =   1815
         MaxLength       =   6
         TabIndex        =   6
         Top             =   4230
         Width           =   830
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   7
         Left            =   1800
         MaxLength       =   6
         TabIndex        =   7
         Top             =   4620
         Width           =   830
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   6
         Left            =   2715
         Locked          =   -1  'True
         TabIndex        =   23
         Text            =   "Text5"
         Top             =   4230
         Width           =   3705
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   7
         Left            =   2715
         Locked          =   -1  'True
         TabIndex        =   22
         Text            =   "Text5"
         Top             =   4605
         Width           =   3705
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   4
         Left            =   1845
         MaxLength       =   3
         TabIndex        =   4
         Tag             =   "Código Postal|T|S|||clientes|codposta|||"
         Top             =   3240
         Width           =   495
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   5
         Left            =   1845
         MaxLength       =   3
         TabIndex        =   5
         Tag             =   "Código Postal|T|S|||clientes|codposta|||"
         Top             =   3600
         Width           =   495
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   3
         Left            =   1845
         MaxLength       =   10
         TabIndex        =   3
         Tag             =   "Código Postal|T|S|||clientes|codposta|||"
         Top             =   2400
         Width           =   1050
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   2
         Left            =   1845
         MaxLength       =   10
         TabIndex        =   2
         Tag             =   "Código Postal|T|S|||clientes|codposta|||"
         Top             =   2040
         Width           =   1050
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   5445
         TabIndex        =   9
         Top             =   5430
         Width           =   975
      End
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   4260
         TabIndex        =   8
         Top             =   5430
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
         Top             =   1215
         Width           =   830
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   2745
         Locked          =   -1  'True
         TabIndex        =   12
         Text            =   "Text5"
         Top             =   840
         Width           =   3555
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   2745
         Locked          =   -1  'True
         TabIndex        =   11
         Text            =   "Text5"
         Top             =   1215
         Width           =   3555
      End
      Begin MSComctlLib.ProgressBar Pb1 
         Height          =   255
         Left            =   570
         TabIndex        =   27
         Top             =   5040
         Visible         =   0   'False
         Width           =   5880
         _ExtentX        =   10372
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   570
         Top             =   5310
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   5
         Left            =   1530
         MouseIcon       =   "frmCentimoSanitario.frx":000C
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar colectivo"
         Top             =   4620
         Width           =   240
      End
      Begin VB.Label Label4 
         Caption         =   "Desde"
         Height          =   195
         Index           =   5
         Left            =   930
         TabIndex        =   26
         Top             =   4230
         Width           =   465
      End
      Begin VB.Label Label4 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   4
         Left            =   930
         TabIndex        =   25
         Top             =   4605
         Width           =   420
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
         Left            =   570
         TabIndex        =   24
         Top             =   3990
         Width           =   660
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   4
         Left            =   1530
         MouseIcon       =   "frmCentimoSanitario.frx":015E
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar colectivo"
         Top             =   4230
         Width           =   240
      End
      Begin VB.Label Label4 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   2
         Left            =   960
         TabIndex        =   21
         Top             =   3600
         Width           =   420
      End
      Begin VB.Label Label4 
         Caption         =   "Desde"
         Height          =   195
         Index           =   1
         Left            =   960
         TabIndex        =   20
         Top             =   3240
         Width           =   465
      End
      Begin VB.Label Label4 
         Caption         =   "Serie Facturas"
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
         Index           =   0
         Left            =   600
         TabIndex        =   19
         Top             =   3000
         Width           =   1815
      End
      Begin VB.Label Label4 
         Caption         =   "Fecha factura"
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
         Index           =   16
         Left            =   600
         TabIndex        =   18
         Top             =   1800
         Width           =   1815
      End
      Begin VB.Label Label4 
         Caption         =   "Desde"
         Height          =   195
         Index           =   15
         Left            =   960
         TabIndex        =   17
         Top             =   2040
         Width           =   465
      End
      Begin VB.Label Label4 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   14
         Left            =   960
         TabIndex        =   16
         Top             =   2400
         Width           =   420
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   2
         Left            =   1530
         Picture         =   "frmCentimoSanitario.frx":02B0
         ToolTipText     =   "Buscar fecha"
         Top             =   2040
         Width           =   240
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   3
         Left            =   1530
         Picture         =   "frmCentimoSanitario.frx":033B
         ToolTipText     =   "Buscar fecha"
         Top             =   2400
         Width           =   240
      End
      Begin VB.Label Label4 
         Caption         =   "Desde"
         Height          =   195
         Index           =   13
         Left            =   960
         TabIndex        =   15
         Top             =   840
         Width           =   465
      End
      Begin VB.Label Label4 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   12
         Left            =   960
         TabIndex        =   14
         Top             =   1215
         Width           =   420
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Cliente"
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
         Left            =   600
         TabIndex        =   13
         Top             =   600
         Width           =   495
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   0
         Left            =   1560
         MouseIcon       =   "frmCentimoSanitario.frx":03C6
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar cliente"
         Top             =   840
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   1
         Left            =   1545
         MouseIcon       =   "frmCentimoSanitario.frx":0518
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar cliente"
         Top             =   1215
         Width           =   240
      End
   End
End
Attribute VB_Name = "frmCentimoSanitario"
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

Public Ajenas As Byte  '1= Ajenas (schfacr)
                       '0= No ajenas (schfac)
    
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


Private Function CargarTemporal() As Boolean
Dim Sql As String

    On Error GoTo eCargarTemporal

    CargarTemporal = False

    Sql = "delete from tmpinformes where codusu = " & vSesion.Codigo
    Conn.Execute Sql
        
    Sql = "insert into tmpinformes (codusu, importe1, nombre1, fecha1, importe2, importe3, importe4)     "
    
'" CASE rcalidad.tipcalid WHEN 0 THEN ""Normal"" WHEN 1 THEN ""Destrio"" WHEN 2 THEN ""Venta Campo"" WHEN 3 THEN ""Mermas"" WHEN 4 THEN ""Pequeño"" END, "
    
    
    Sql = Sql & "select " & vSesion.Codigo & ", schfacr.codsocio, concat(right(concat('0000000',schfacr.numfactu),7),'-',schfacr.letraser) numfactu, schfacr.fecfactu, "
    Sql = Sql & "sum(slhfacr.implinea) importe,  CASE sartic.tipogaso WHEN 1 THEN 1 WHEN 2 THEN 2 WHEN 3 THEN 3 WHEN 4 THEN 3 END codartic,  sum(slhfacr.cantidad) cantidad "
    Sql = Sql & " from (((schfacr  inner join slhfacr on schfacr.letraser = slhfacr.letraser and schfacr.numfactu = slhfacr.numfactu and "
    Sql = Sql & " schfacr.fecfactu = slhfacr.fecfactu) "
    Sql = Sql & " inner join sartic on sartic.codartic = slhfacr.codartic) "
    Sql = Sql & " inner join ssocio on ssocio.codsocio = schfacr.codsocio) "
    Sql = Sql & " inner join sfamia on sartic.codfamia = sfamia.codfamia and sfamia.tipfamia = 1"
    Sql = Sql & " where sartic.tipogaso between 1 and 4 "
    
    If txtCodigo(0).Text <> "" Then Sql = Sql & " and schfacr.codsocio >= " & DBSet(txtCodigo(0).Text, "N")
    If txtCodigo(1).Text <> "" Then Sql = Sql & " and schfacr.codsocio <= " & DBSet(txtCodigo(1).Text, "N")
    
    If txtCodigo(2).Text <> "" Then Sql = Sql & " and schfacr.fecfactu >= " & DBSet(txtCodigo(2).Text, "F")
    If txtCodigo(3).Text <> "" Then Sql = Sql & " and schfacr.fecfactu <= " & DBSet(txtCodigo(3).Text, "F")

    If txtCodigo(4).Text <> "" Then Sql = Sql & " and schfacr.letraser >= " & DBSet(txtCodigo(4).Text, "T")
    If txtCodigo(5).Text <> "" Then Sql = Sql & " and schfacr.letraser <= " & DBSet(txtCodigo(5).Text, "T")

    If txtCodigo(6).Text <> "" Then Sql = Sql & " and ssocio.codcoope >= " & DBSet(txtCodigo(6).Text, "N")
    If txtCodigo(7).Text <> "" Then Sql = Sql & " and ssocio.codcoope <= " & DBSet(txtCodigo(7).Text, "N")


    Sql = Sql & " group by 1,2,3,4,6 "
    Sql = Sql & " order by 1,2,3"

    ' en caso de no ser ajenas cambiamos las tablas a las normales
    If Ajenas = 0 Then
        Sql = Replace(Replace(Sql, "schfacr", "schfac"), "slhfacr", "slhfac")
    End If


    Conn.Execute Sql
    
    CargarTemporal = True
    Exit Function
    
eCargarTemporal:
    MuestraError Err.Number, "Cargar Temporal", Err.Description
End Function


Private Sub cmdAceptar_Click()
'Listado cobros por fecha de vencimiento
Dim cDesde As String, cHasta As String 'cadena codigo Desde/Hasta
Dim nDesde As String, nHasta As String 'cadena Descripcion Desde/Hasta
Dim cadTABLA As String, cOrden As String
Dim i As Byte
Dim NRegs As Long
Dim b As Boolean
Dim NomFic As String
                 
    InicializarVbles
    
    cadParam = cadParam & "|pEmpresa=""" & vEmpresa.nomEmpre & """|"
    numParam = numParam + 1
    
    
    'D/H Cliente
    cDesde = Trim(txtCodigo(0).Text)
    cHasta = Trim(txtCodigo(1).Text)
    nDesde = txtNombre(0).Text
    nHasta = txtNombre(1).Text
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{" & tabla & ".codsocio}"
        TipCod = "N"
        If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHcliente= """) Then Exit Sub
    End If
    
    'D/H Fecha factura
    cDesde = Trim(txtCodigo(2).Text)
    cHasta = Trim(txtCodigo(3).Text)
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{" & tabla & ".fecfactu}"
        TipCod = "F"
        If Not PonerDesdeHasta(cDesde, cHasta, "", "", "pDHfechaFac= """) Then Exit Sub
    End If
    
    'D/H Colectivo
    cDesde = Trim(txtCodigo(6).Text)
    cHasta = Trim(txtCodigo(7).Text)
    nDesde = txtNombre(6).Text
    nHasta = txtNombre(7).Text
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{ssocio.codcoope}"
        TipCod = "N"
        If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHColec= """) Then Exit Sub
    End If
    
    
    
    If CargarTemporal Then
    
        NRegs = DevuelveValor("select count(*) from tmpinformes where codusu = " & vSesion.Codigo)
        If NRegs <> 0 Then
            
            If Not PrepararCarpetas(False) Then Exit Sub
            
            NomFic = GetFolder("Carpeta de Socios")
            If NomFic = "" Then Exit Sub
            
            Pb1.visible = True
            Pb1.Max = NRegs
            Pb1.Value = 0
                
            b = GeneraFichero(Pb1, NomFic)
            
            If b Then
                 MsgBox "Proceso realizado correctamente", vbExclamation
                
                 
                 cadFormula = "{tmpinformes.codusu} = " & vSesion.Codigo
                 
                 cadTitulo = "Centimo Sanitario"
                 cadNombreRPT = "rCentimoSanitario.rpt"
                 
                 LlamarImprimir
                
                 Pb1.visible = False
            End If
        Else
            MsgBox "No hay registros para generar el fichero", vbExclamation
        End If
    End If
End Sub


Private Function GeneraFichero(ByRef Pb1 As ProgressBar, NomDir As String) As Boolean
Dim NFich As Integer
Dim Rs As ADODB.Recordset
Dim Sql As String

Dim SocioAnt As Long
Dim Linea As Long
Dim cad As String

Dim Sql3 As String
Dim Rs3 As ADODB.Recordset
Dim NomFichero As String

    On Error GoTo EGen
    
    GeneraFichero = False

    Sql = "select importe1, nombre1, fecha1, importe2, importe3, importe4 from tmpinformes where codusu = " & vSesion.Codigo
    Sql = Sql & " order by 1, 3, 2, 5 "
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    If Not Rs.EOF Then
        SocioAnt = Rs!Importe1
    
        Sql3 = "select nrosocio, nomsocio from ssocio where codsocio = " & DBSet(SocioAnt, "N")
        Set Rs3 = New ADODB.Recordset
        Rs3.Open Sql3, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        NomFichero = DBLet(Rs3!NomSocio, "T")
        NomFichero = Trim(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(NomFichero, "/", ""), "\", ""), ":", ""), "?", ""), "*", ""), "<", ""), ">", ""), "|", ""), """", ""))
        If DBLet(Rs3!nrosocio, "N") <> 0 Then NomFichero = NomFichero & " (" & Format(DBLet(Rs3!nrosocio, "N"), "000000") & ")"
        NomFichero = NomFichero & ".txt"
        
        NFich = FreeFile
        Open NomDir & "\" & NomFichero For Output As #NFich
    End If
    
    Linea = 0
    While Not Rs.EOF
        If SocioAnt <> Rs!Importe1 Then
            Close (NFich)
            
            SocioAnt = Rs!Importe1
        
            Sql3 = "select nrosocio, nomsocio from ssocio where codsocio = " & DBSet(SocioAnt, "N")
            Set Rs3 = New ADODB.Recordset
            Rs3.Open Sql3, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            NomFichero = DBLet(Rs3!NomSocio, "T")
            NomFichero = Trim(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(NomFichero, "/", ""), "\", ""), ":", ""), "?", ""), "*", ""), "<", ""), ">", ""), "|", ""), """", ""))
            If DBLet(Rs3!nrosocio, "N") <> 0 Then NomFichero = NomFichero & " (" & Format(DBLet(Rs3!nrosocio, "N"), "000000") & ")"
            NomFichero = NomFichero & ".txt"
        
            NFich = FreeFile
            Open NomDir & "\" & NomFichero For Output As #NFich
        End If
        
'        IncrementarProgres Pb1, 1
        Pb1.Value = Pb1.Value + 1
        
        Linea = Linea + 1
        
        cad = "17;" 'comunidad valenciana
        cad = cad & Trim(Rs!nombre1) & ";" ' factura
        cad = cad & Format(DBLet(Rs!Fecha1, "F"), "dd-mm-yyyy") & ";"
        cad = cad & Trim(Format(DBLet(Rs!Importe2, "N"), "########0.00")) & ";"
        cad = cad & DBLet(Rs!importe3, "N") & ";"
        cad = cad & Trim(Format(DBLet(Rs!Importe4, "N"), "########0.00")) & ";"
        
        Print #NFich, cad
        
        Rs.MoveNext
    Wend
    

    Close (NFich)
    If Linea > 0 Then GeneraFichero = True
    Exit Function

EGen:
    Set Rs = Nothing
    Close (NFich)
    MuestraError Err.Number, Err.Description
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
     Me.imgBuscar(4).Picture = frmPpal.imgListImages16.ListImages(1).Picture
     Me.imgBuscar(5).Picture = frmPpal.imgListImages16.ListImages(1).Picture

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
        
        Case 4, 5 'COLECTIVO
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
'15/02/2007
'    KEYpress KeyAscii
' ahora
    If KeyAscii = teclaBuscar Then
        Select Case Index
            Case 0: KEYBusqueda KeyAscii, 0 'cliente desde
            Case 1: KEYBusqueda KeyAscii, 1 'cliente hasta
            Case 2: KEYFecha KeyAscii, 2 'fecha desde
            Case 3: KEYFecha KeyAscii, 3 'fecha hasta
            Case 6: KEYBusqueda KeyAscii, 4 'colectivo desde
            Case 7: KEYBusqueda KeyAscii, 5 'colectivo hasta
            
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
        
        Case 2, 3 'FECHAS
            If txtCodigo(Index).Text <> "" Then PonerFormatoFecha txtCodigo(Index)
            
        Case 4, 5 'SERIE
              txtCodigo(Index).Text = UCase(txtCodigo(Index).Text)
              
        Case 6, 7 'COLECTIVO
            txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), "scoope", "nomcoope", "codcoope", "N")
            If txtCodigo(Index).Text <> "" Then txtCodigo(Index).Text = Format(txtCodigo(Index).Text, "00")
            
              
  End Select
End Sub

Private Sub FrameCobrosVisible(visible As Boolean, ByRef h As Integer, ByRef w As Integer)
'Frame para los cobros a clientes por fecha vencimiento
    Me.FrameCobros.visible = visible
    If visible = True Then
        Me.FrameCobros.Top = -90
        Me.FrameCobros.Left = 0
        Me.FrameCobros.Height = 6735
        Me.FrameCobros.Width = 7275
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
 
Private Sub AbrirFrmColectivo(indice As Integer)
    indCodigo = indice + 2
    Set frmCol = New frmManCoope
    frmCol.DatosADevolverBusqueda = "0|1|"
    frmCol.DeConsulta = True
    frmCol.CodigoActual = txtCodigo(indCodigo)
    frmCol.Show vbModal
    Set frmCol = Nothing
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


Private Function PrepararCarpetas(Optional NoBorrar As Boolean) As Boolean
    On Error GoTo EPrepararCarpetas
    
    PrepararCarpetas = False

    If Dir(App.path & "\temp", vbDirectory) = "" Then
        MkDir App.path & "\temp"
    Else
        If Not NoBorrar Then
            If Dir(App.path & "\temp\*.*", vbArchive) <> "" Then Kill App.path & "\temp\*.*"
        End If
    End If

    PrepararCarpetas = True
    Exit Function
    
EPrepararCarpetas:
    MuestraError Err.Number, "", "Preparar Carpetas"
End Function

