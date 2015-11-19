VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmEstFPago 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Estadística Forma Pago / Turno"
   ClientHeight    =   5805
   ClientLeft      =   45
   ClientTop       =   2430
   ClientWidth     =   6465
   Icon            =   "frmEstFpago.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5805
   ScaleWidth      =   6465
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
      Height          =   5835
      Left            =   0
      TabIndex        =   10
      Top             =   -60
      Width           =   6555
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   7
         Left            =   1830
         MaxLength       =   3
         TabIndex        =   7
         Top             =   4350
         Width           =   830
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   6
         Left            =   1830
         MaxLength       =   3
         TabIndex        =   6
         Top             =   3990
         Width           =   830
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   6
         Left            =   2715
         Locked          =   -1  'True
         TabIndex        =   27
         Text            =   "Text5"
         Top             =   3990
         Width           =   3135
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   7
         Left            =   2715
         Locked          =   -1  'True
         TabIndex        =   26
         Text            =   "Text5"
         Top             =   4365
         Width           =   3135
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   5
         Left            =   2745
         Locked          =   -1  'True
         TabIndex        =   20
         Text            =   "Text5"
         Top             =   3375
         Width           =   3135
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   4
         Left            =   2745
         Locked          =   -1  'True
         TabIndex        =   19
         Text            =   "Text5"
         Top             =   3000
         Width           =   3135
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   5
         Left            =   1845
         MaxLength       =   2
         TabIndex        =   5
         Top             =   3375
         Width           =   830
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   4
         Left            =   1845
         MaxLength       =   2
         TabIndex        =   4
         Top             =   3000
         Width           =   830
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   3
         Left            =   1845
         MaxLength       =   10
         TabIndex        =   3
         Tag             =   "Código Postal|T|S|||clientes|codposta|||"
         Top             =   2280
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
         Top             =   1920
         Width           =   1050
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   4890
         TabIndex        =   9
         Top             =   5220
         Width           =   975
      End
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   3690
         TabIndex        =   8
         Top             =   5220
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
         Width           =   3135
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
         Width           =   3135
      End
      Begin MSComctlLib.ProgressBar Pb1 
         Height          =   255
         Left            =   540
         TabIndex        =   24
         Top             =   4830
         Visible         =   0   'False
         Width           =   5325
         _ExtentX        =   9393
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Label Label4 
         Caption         =   "Desde"
         Height          =   195
         Index           =   6
         Left            =   930
         TabIndex        =   30
         Top             =   3990
         Width           =   465
      End
      Begin VB.Label Label4 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   5
         Left            =   930
         TabIndex        =   29
         Top             =   4365
         Width           =   420
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Familia"
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
         Index           =   4
         Left            =   570
         TabIndex        =   28
         Top             =   3750
         Width           =   480
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   3
         Left            =   1530
         MouseIcon       =   "frmEstFpago.frx":000C
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar Familia"
         Top             =   4380
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   2
         Left            =   1530
         MouseIcon       =   "frmEstFpago.frx":015E
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar Familia"
         Top             =   3990
         Width           =   240
      End
      Begin VB.Label Label4 
         Caption         =   "Cargando Tabla Temporal"
         Height          =   195
         Index           =   3
         Left            =   540
         TabIndex        =   25
         Top             =   5100
         Visible         =   0   'False
         Width           =   1845
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   5
         Left            =   1545
         MouseIcon       =   "frmEstFpago.frx":02B0
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar F.Pago"
         Top             =   3375
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   4
         Left            =   1560
         MouseIcon       =   "frmEstFpago.frx":0402
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar F.Pago"
         Top             =   3000
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Forma de Pago"
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
         Left            =   600
         TabIndex        =   23
         Top             =   2760
         Width           =   1080
      End
      Begin VB.Label Label4 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   1
         Left            =   960
         TabIndex        =   22
         Top             =   3375
         Width           =   420
      End
      Begin VB.Label Label4 
         Caption         =   "Desde"
         Height          =   195
         Index           =   0
         Left            =   960
         TabIndex        =   21
         Top             =   3000
         Width           =   465
      End
      Begin VB.Label Label4 
         Caption         =   "Fecha albarán"
         ForeColor       =   &H00972E0B&
         Height          =   255
         Index           =   16
         Left            =   600
         TabIndex        =   18
         Top             =   1680
         Width           =   1815
      End
      Begin VB.Label Label4 
         Caption         =   "Desde"
         Height          =   195
         Index           =   15
         Left            =   960
         TabIndex        =   17
         Top             =   1920
         Width           =   465
      End
      Begin VB.Label Label4 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   14
         Left            =   960
         TabIndex        =   16
         Top             =   2280
         Width           =   420
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   2
         Left            =   1530
         Picture         =   "frmEstFpago.frx":0554
         ToolTipText     =   "Buscar fecha"
         Top             =   1920
         Width           =   240
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   3
         Left            =   1530
         Picture         =   "frmEstFpago.frx":05DF
         ToolTipText     =   "Buscar fecha"
         Top             =   2280
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
         MouseIcon       =   "frmEstFpago.frx":066A
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar Cliente"
         Top             =   840
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   1
         Left            =   1545
         MouseIcon       =   "frmEstFpago.frx":07BC
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar Cliente"
         Top             =   1215
         Width           =   240
      End
   End
End
Attribute VB_Name = "frmEstFPago"
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
Private WithEvents frmFam As frmManFamia 'Familias
Attribute frmFam.VB_VarHelpID = -1
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

Dim CFPago(11) As String
Dim NFPago(11) As String

Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, 0, cerrar
    If cerrar Then Unload Me
End Sub

Private Sub cmdAceptar_Click()
Dim cDesde As String, cHasta As String 'cadena codigo Desde/Hasta
Dim nDesde As String, nHasta As String 'cadena Descripcion Desde/Hasta
Dim cadTABLA As String, cOrden As String
Dim I As Byte

Dim Sql As String
Dim Rs As ADODB.Recordset

    InicializarVbles
    
    '========= PARAMETROS  =============================
    'Añadir el parametro de Empresa
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
        Codigo = "{" & tabla & ".fecalbar}"
        TipCod = "F"
        If Not PonerDesdeHasta(cDesde, cHasta, "", "", "pDHfechaFac= """) Then Exit Sub
    End If
    
    
    'D/H Familia
    cDesde = Trim(txtCodigo(6).Text)
    cHasta = Trim(txtCodigo(7).Text)
    nDesde = txtNombre(6).Text
    nHasta = txtNombre(7).Text
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{sartic.codfamia}"
        TipCod = "N"
        If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHFamilia= """) Then Exit Sub
    End If
    
    
    
    'D/H F.Pago
    cDesde = Trim(txtCodigo(4).Text)
    cHasta = Trim(txtCodigo(5).Text)
    nDesde = txtNombre(4).Text
    nHasta = txtNombre(5).Text
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{scaalb.codforpa}"
        TipCod = "N"
        If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHFpago= """) Then Exit Sub
    End If
    
    
    For I = 0 To 10
        CFPago(I) = -1
        NFPago(I) = ""
    Next I
    
    Sql = "select codforpa, nomforpa from sforpa where 1=1  "
    If cDesde <> "" Then Sql = Sql & " and codforpa >= " & DBSet(cDesde, "N")
    If cHasta <> "" Then Sql = Sql & " and codforpa <= " & DBSet(cHasta, "N")
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    'cargamos las formas de pago
    I = 0
    While Not Rs.EOF And I <= 10
        CFPago(I) = DBLet(Rs!Codforpa, "N")
        NFPago(I) = DBLet(Rs!nomforpa, "T")
        
        cadParam = cadParam & "pNFPago" & I & "=""" & NFPago(I) & """|"
        numParam = numParam + 1
        
        I = I + 1
        Rs.MoveNext
    Wend
    
    cadParam = cadParam & "pCount=" & I & "|"
    numParam = numParam + 1
    
    'Comprobar si hay registros a Mostrar antes de abrir el Informe
    cadTABLA = tabla & " INNER JOIN ssocio ON " & tabla & ".codsocio=ssocio.codsocio "
    cadTABLA = "(" & cadTABLA & ") INNER JOIN sartic ON scaalb.codartic = sartic.codartic "
    
    If HayRegParaInforme(cadTABLA, cadSelect) Then
        If CargarTemporal(cadTABLA, cadSelect) Then
            cadFormula = "{tmpinformes.codusu} = " & vSesion.Codigo
            
            cadTitulo = "Estadística Forma de Pago / Turno"
            cadNombreRPT = "rEstFPago.rpt"
            LlamarImprimir
        End If
    End If
End Sub


Private Function CargarTemporal(cTabla As String, cWhere As String) As Boolean
Dim Sql As String
Dim Rs As ADODB.Recordset
Dim I As Integer

Dim F0 As Currency
Dim F1 As Currency
Dim F2 As Currency
Dim F3 As Currency
Dim F4 As Currency
Dim F5 As Currency
Dim F6 As Currency
Dim F7 As Currency
Dim F8 As Currency
Dim F9 As Currency
Dim F10 As Currency
Dim NRegs As Integer


    On Error GoTo eCargarTemporal
    
    CargarTemporal = False
    
    Sql = "delete from tmpinformes where codusu = " & vSesion.Codigo
    Conn.Execute Sql
    
    Sql = "insert into tmpinformes (codusu, fecha1, campo1, importe1, importe2, importe3, importe4, importe5, importe6, "
    Sql = Sql & "importeb1, importeb2, importeb3, importeb4, importeb5) select distinct " & vSesion.Codigo
    Sql = Sql & ", fecalbar,sartic.codfamia,0,0,0,0,0,0,0,0,0,0,0 from " & cTabla
    If cWhere <> "" Then Sql = Sql & " where " & cWhere
    
    Conn.Execute Sql
    
    
    Sql = "select * from tmpinformes where codusu = " & vSesion.Codigo
    Sql = Sql & " order by codigo1, fecha1 "
    
    NRegs = TotalRegistrosConsulta(Sql)
    
    CargarProgres Pb1, CInt(NRegs)
    Pb1.visible = True
    Label4(3).visible = True
    DoEvents
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    While Not Rs.EOF
        IncrementarProgres Pb1, 1
        DoEvents
    
        F0 = 0
        F1 = 0
        F2 = 0
        F3 = 0
        F4 = 0
        F5 = 0
        F6 = 0
        F7 = 0
        F8 = 0
        F9 = 0
        F10 = 0
        
        For I = 0 To 10
            If CFPago(I) <> -1 Then
                Sql = "select sum(importel) from " & cTabla & " where fecalbar = " & DBSet(Rs!Fecha1, "F") & " and scaalb.codforpa = " & DBSet(CFPago(I), "N") & " and sartic.codfamia = " & DBSet(Rs!campo1, "N")
                If cWhere <> "" Then Sql = Sql & " and " & cWhere
            
                Select Case I
                    Case 0
                        F0 = DevuelveValor(Sql)
                    Case 1
                        F1 = DevuelveValor(Sql)
                    Case 2
                        F2 = DevuelveValor(Sql)
                    Case 3
                        F3 = DevuelveValor(Sql)
                    Case 4
                        F4 = DevuelveValor(Sql)
                    Case 5
                        F5 = DevuelveValor(Sql)
                    Case 6
                        F6 = DevuelveValor(Sql)
                    Case 7
                        F7 = DevuelveValor(Sql)
                    Case 8
                        F8 = DevuelveValor(Sql)
                    Case 9
                        F9 = DevuelveValor(Sql)
                    Case 10
                        F10 = DevuelveValor(Sql)
                End Select
            End If
        Next I
        
        ' actualizamos la temporal
        Sql = "update tmpinformes set "
        Sql = Sql & " importe1 = " & DBSet(F0, "N")
        Sql = Sql & ",importe2 = " & DBSet(F1, "N")
        Sql = Sql & ",importe3 = " & DBSet(F2, "N")
        Sql = Sql & ",importe4 = " & DBSet(F3, "N")
        Sql = Sql & ",importe5 = " & DBSet(F4, "N")
        Sql = Sql & ",importe6 = " & DBSet(F5, "N")
        Sql = Sql & ",importeb1 = " & DBSet(F6, "N")
        Sql = Sql & ",importeb2 = " & DBSet(F7, "N")
        Sql = Sql & ",importeb3 = " & DBSet(F8, "N")
        Sql = Sql & ",importeb4 = " & DBSet(F9, "N")
        Sql = Sql & ",importeb5 = " & DBSet(F10, "N")
        Sql = Sql & " where codusu = " & vSesion.Codigo
        Sql = Sql & " and fecha1 = " & DBSet(Rs!Fecha1, "F")
        Sql = Sql & " and campo1 = " & DBSet(Rs!campo1, "N")
        
        Conn.Execute Sql
        
        Rs.MoveNext
    Wend
    Set Rs = Nothing
    
    CargarTemporal = True
    Pb1.visible = False
    Label4(3).visible = False
    DoEvents
    
    Exit Function

eCargarTemporal:
    MuestraError Err.Number, "Cargando Temporal", Err.Description
    Pb1.visible = False
    Label4(3).visible = False
    DoEvents
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
Dim I As Integer


    PrimeraVez = True
    limpiar Me

    'IMAGES para busqueda
    For I = 0 To 5
        Me.imgBuscar(I).Picture = frmPpal.imgListImages16.ListImages(1).Picture
    Next I

    '###Descomentar
'    CommitConexion
         
    FrameCobrosVisible True, h, w
    indFrame = 5
    tabla = "scaalb"
            
    'Esto se consigue poneinedo el cancel en el opcion k corresponda
    Me.cmdCancel.Cancel = True
    
'   Me.Width = w + 70
'   Me.Height = h + 350
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

Private Sub frmFam_DatoSeleccionado(CadenaSeleccion As String)
'Form de Consulta de Familias
    txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmFpa_DatoSeleccionado(CadenaSeleccion As String)
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
        
        Case 4, 5 'F.PAGO
            AbrirFrmFpagos (Index)
        
        Case 2, 3 ' Familia
            AbrirFrmFamilias (Index)
        
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
    If KeyAscii = teclaBuscar Then
        Select Case Index
            Case 0: KEYBusqueda KeyAscii, 0 'cliente desde
            Case 1: KEYBusqueda KeyAscii, 1 'cliente hasta
            Case 4: KEYBusqueda KeyAscii, 4 'forma de pago desde
            Case 5: KEYBusqueda KeyAscii, 5 'forma de pago hasta
            Case 2: KEYFecha KeyAscii, 2 'fecha albaran desde
            Case 3: KEYFecha KeyAscii, 3 'fecha albaran hasta
            Case 6: KEYBusqueda KeyAscii, 2 'familia desde
            Case 7: KEYBusqueda KeyAscii, 3 'familia hasta
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
            
        Case 4, 5 'F.PAGO
            txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), "sforpa", "nomforpa", "codforpa", "N")
            If txtCodigo(Index).Text <> "" Then txtCodigo(Index).Text = Format(txtCodigo(Index).Text, "00")
        
        Case 6, 7 'Familia
            txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), "sfamia", "nomfamia", "codfamia", "N")
            If txtCodigo(Index).Text <> "" Then txtCodigo(Index).Text = Format(txtCodigo(Index).Text, "000")
            
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

Private Sub AbrirFrmFpagos(indice As Integer)
    indCodigo = indice
    Set frmFPa = New frmManFpago
    frmFPa.DatosADevolverBusqueda = "0|1|"
    frmFPa.DeConsulta = True
    frmFPa.CodigoActual = txtCodigo(indCodigo)
    frmFPa.Show vbModal
    Set frmFPa = Nothing
End Sub
 
Private Sub AbrirFrmFamilias(indice As Integer)
    indCodigo = indice + 4
    Set frmFam = New frmManFamia
    frmFam.DatosADevolverBusqueda = "0|1|"
    frmFam.DeConsulta = True
    frmFam.CodigoActual = txtCodigo(indCodigo)
    frmFam.Show vbModal
    Set frmFam = Nothing
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

