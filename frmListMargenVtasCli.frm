VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmListMargenVtasCli 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Margen de Ventas por Cliente"
   ClientHeight    =   4020
   ClientLeft      =   45
   ClientTop       =   2430
   ClientWidth     =   8040
   Icon            =   "frmListMargenVtasCli.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4020
   ScaleWidth      =   8040
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameMovArtic 
      Height          =   4035
      Left            =   0
      TabIndex        =   6
      Top             =   -30
      Width           =   7995
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   2
         Left            =   1860
         MaxLength       =   10
         TabIndex        =   2
         Tag             =   "Código Postal|T|S|||clientes|codposta|||"
         Top             =   1800
         Width           =   1050
      End
      Begin VB.CheckBox chkDetalla 
         Caption         =   "Resumen"
         Height          =   255
         Left            =   630
         TabIndex        =   3
         Top             =   2400
         Width           =   2055
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   3
         Left            =   6180
         TabIndex        =   5
         Top             =   3270
         Width           =   975
      End
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "&Aceptar"
         Height          =   375
         Index           =   3
         Left            =   5100
         TabIndex        =   4
         Top             =   3270
         Width           =   975
      End
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Index           =   16
         Left            =   1860
         MaxLength       =   4
         TabIndex        =   0
         Tag             =   "Código Postal|T|S|||clientes|codposta|00||"
         Top             =   1290
         Width           =   615
      End
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Index           =   17
         Left            =   2670
         MaxLength       =   4
         TabIndex        =   1
         Tag             =   "Código Postal|T|S|||clientes|codposta|0000||"
         Top             =   1290
         Width           =   615
      End
      Begin MSComctlLib.ProgressBar Pb1 
         Height          =   255
         Left            =   600
         TabIndex        =   10
         Top             =   2820
         Visible         =   0   'False
         Width           =   6705
         _ExtentX        =   11827
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Label Label4 
         Caption         =   "Cargando Tabla Temporal"
         Height          =   195
         Index           =   0
         Left            =   600
         TabIndex        =   11
         Top             =   3090
         Visible         =   0   'False
         Width           =   4365
      End
      Begin VB.Label Label4 
         Caption         =   "Precio de Coste"
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
         Left            =   660
         TabIndex        =   9
         Top             =   1830
         Width           =   1245
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Mes / Año"
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
         Left            =   660
         TabIndex        =   7
         Top             =   1320
         Width           =   720
      End
      Begin VB.Label lbltituloInven 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   345
         Left            =   570
         TabIndex        =   8
         Top             =   330
         Width           =   6945
      End
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   7110
      Top             =   3540
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox txtCodigo 
      Height          =   285
      Index           =   14
      Left            =   1380
      MaxLength       =   16
      TabIndex        =   12
      Top             =   420
      Width           =   1455
   End
   Begin VB.TextBox txtNombre 
      BackColor       =   &H80000018&
      Height          =   285
      Index           =   14
      Left            =   2895
      Locked          =   -1  'True
      TabIndex        =   13
      Text            =   "Text5"
      Top             =   420
      Width           =   3855
   End
   Begin VB.Image imgBuscarG 
      Height          =   240
      Index           =   11
      Left            =   990
      Top             =   420
      Width           =   240
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Artículo"
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
      Left            =   180
      TabIndex        =   14
      Top             =   420
      Width           =   540
   End
End
Attribute VB_Name = "frmListMargenVtasCli"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public OpcionListado As Integer

    '==== Listados de ALMACEN ====
    '=============================
    ' 9 .- Listado Busquedas de movimientos de Artículos
    '10 .-
    
    
Public NumCod As String 'Para indicar cod. Traspaso,Movimiento, etc. que llama
                        'Para indicar nº oferta a imprimir

Public CadTag As String 'Cadena con el Tag del campo que se va a poner en D/H en los listados
                        'Se necesita si el tipo de codigo es texto

    
Private Conexion As Byte
'1.- Conexión a BD Ariges  2.- Conexión a BD Conta

Private HaDevueltoDatos As Boolean

Private WithEvents frmMtoArticulos As frmManArtic
Attribute frmMtoArticulos.VB_VarHelpID = -1
Private WithEvents frmMtoClientes As frmManClien
Attribute frmMtoClientes.VB_VarHelpID = -1
Private WithEvents frmMens As frmMensajes
Attribute frmMens.VB_VarHelpID = -1
Private WithEvents frmB As frmBuscaGrid  'Busquedas
Attribute frmB.VB_VarHelpID = -1
Private WithEvents frmF As frmCal 'Calendario de Fechas
Attribute frmF.VB_VarHelpID = -1

'---- Variables para el INFORME ----
Private cadFormula As String 'Cadena con la FormulaSelection para Crystal Report
Private cadParam As String 'Cadena con los parametros para Crystal Report
Private numParam As Byte 'Numero de parametros que se pasan a Crystal Report
Private cadSelect As String 'Cadena para comprobar si hay datos antes de abrir Informe
Private cadTitulo As String 'Titulo para el frmImprimir
Private cadNomRPT As String 'Nombre del informe a Imprimir
Private conSubRPT As Boolean 'Si el informe tiene subreports
'-----------------------------------


Dim indCodigo As Integer 'indice para txtCodigo

Dim Codigo As String 'Código para FormulaSelection de Crystal Report
Dim Orden1 As String 'Campo de Ordenacion (por codigo) para Cristal Report
Dim Orden2 As String 'Campo de Ordenacion (por nombre) para Cristal Report

Dim PrimeraVez As Boolean
Dim indFrame As Single

Dim ArtDto As String

Dim Articulos As String
Dim OK As Boolean
Dim NomArtic As String
Dim Articulo As String


Private Sub KEYpress(KeyAscii As Integer)
    Dim cerrar As Boolean

    KEYpressGnral KeyAscii, 2, cerrar
    If cerrar Then Unload Me
End Sub

Private Sub cmdAceptar_Click(Index As Integer)
Dim indRPT As Byte 'Indica el tipo de Documento en la tabla "scryst"
Dim cadAux As String
Dim bol As Boolean

Dim Sql As String

    InicializarVbles
   
    '[Monica]10/11/2014: poder seleccionar mas de un articulo de carburante
    Articulos = ""
    ArtDto = ""
    OK = False
    NomArtic = ""
   
    If Not DatosOk Then Exit Sub
   
    
    '[Monica]10/11/2014: poder seleccionar mas de un articulo de carburante
    Set frmMens = frmMensajes
    
    frmMens.OpcionMensaje = 24
    frmMens.cadWHERE = "sfamia.tipfamia = 1"
    frmMens.cmdEtiqEstan(1).Caption = "&Aceptar"
    frmMens.Show vbModal
    
    Set frmMens = Nothing
   
    
    'Parametro EMPRESA
    cadParam = "|pNomEmpre=""" & vEmpresa.nomEmpre & """|"
    numParam = 1
   
    '[Monica]10/11/2014: poder seleccionar mas de un articulo de carburante
    If Articulos = "" Then
        MsgBox "Debe de introducir al menos un artículo. Revise.", vbExclamation
        Exit Sub
    End If
    
    If Not OK Then
        MsgBox "Los artículos que seleccione deben de tener el mismo artículo de descuento. Revise.", vbExclamation
        Exit Sub
    End If
   
   
'[Monica]10/11/2014: poder seleccionar mas de un articulo de carburante
'   ArtDto = DevuelveValor("select artdto from sartic where codartic = " & DBSet(txtCodigo(14).Text, "N"))
'
'
'   Articulos = "(" & DBSet(txtCodigo(14).Text, "N")
'   If ArtDto <> 0 Then
'        Articulos = Articulos & "," & DBSet(ArtDto, "N") & ")"
'   Else
'        Articulos = Articulos & ")"
'   End If
   
   Sql = "slhfac.codartic in " & Articulos
   Sql = Sql & " and month(slhfac.fecfactu) = " & DBSet(txtCodigo(16).Text, "N") & " and year(slhfac.fecfactu) = " & DBSet(txtCodigo(17).Text, "N")
   
   cadAux = "slhfac"
   
   If CargarTemporal(Sql) Then
        conSubRPT = False
        cadNomRPT = "rFacEstMargenCli.rpt"
                                                             '[Monica]10/11/2014: poder seleccionar mas de un articulo de carburante
        cadParam = cadParam & "pArtic=""" & NomArtic & """|" '& txtNombre(14).Text & """|"
        cadParam = cadParam & "pPrcoste=" & TransformaComasPuntos(txtCodigo(2).Text) & "|"
        cadParam = cadParam & "pMes=" & txtCodigo(16).Text & "|"
        cadParam = cadParam & "pAnyo=" & txtCodigo(17).Text & "|"
        numParam = numParam + 4
        
    
        'Nombre fichero .rpt a Imprimir
        If Me.chkDetalla.Value = 0 Then
            cadParam = cadParam & "pResumen=0|"
        Else
            cadParam = cadParam & "pResumen=1|"
        End If
        numParam = numParam + 1
        
        cadFormula = "{tmpinformes.codusu} = " & vSesion.Codigo
   
        LlamarImprimir
   End If
   
   Me.Pb1.visible = False
   Me.Label4(0).visible = False
   
   Screen.MousePointer = vbDefault
   
End Sub

Private Function CargarTemporal(cWhere As String) As Boolean
Dim Sql As String
Dim Sql2 As String
Dim CadInsert As String
Dim Rs As ADODB.Recordset
Dim PrCoste As Currency
Dim AlbarAnt As String

Dim PrecioAnt As Currency
Dim SocioAnt As Long
Dim ArticAnt As String
Dim LitrosAnt As Currency

Dim FactuAnt As Long

Dim CadValues As String
Dim Importe As Currency
Dim ImpCoste As Currency
Dim PrecioDto As Currency
Dim nRegs As Long
Dim CodIVA As String
Dim vPorcIva As String
Dim PorcIva As Currency
Dim I As Integer
Dim Litros2 As Currency
Dim Sql22 As String
Dim Importe2 As Currency
Dim Fecha1 As Date

    On Error GoTo eCargarTemporal
    
    CargarTemporal = False
   
    Screen.MousePointer = vbHourglass
    
    If Not HayRegParaInforme("slhfac", cWhere) Then Exit Function
    
    
    PrCoste = ImporteSinFormato(txtCodigo(2).Text)
    CodIVA = DevuelveDesdeBDNew(cPTours, "sartic", "codigiva", "codartic", Articulo, "N") ' txtCodigo(14).Text
    
    '[Monica]19/09/2013: fallaba cuando el cambio
    Fecha1 = CDate("01/" & Format(txtCodigo(16).Text, "00") & "/" & Format(txtCodigo(17).Text, "0000"))
    If Fecha1 < vParamAplic.FechaCamIva Then
        Select Case CodIVA
            Case vParamAplic.CodIvaGnral
                CodIVA = vParamAplic.CodIvaGnralAnt
            Case vParamAplic.CodIvaRedu
                CodIVA = vParamAplic.CodIvaReduAnt
        End Select
    End If
    
    vPorcIva = DevuelveDesdeBDNew(cConta, "tiposiva", "porceiva", "codigiva", CodIVA, "N")
    If vPorcIva = "" Then
        PorcIva = 0
    Else
        PorcIva = CCur(vPorcIva)
    End If
    
    'borramos de la temporal
    Sql = "delete from tmpslhfac where codusu = " & vSesion.Codigo
    Conn.Execute Sql
    
    Sql = "delete from tmpinformes where codusu = " & vSesion.Codigo
    Conn.Execute Sql
    
    Me.Label4(0).visible = True
    Me.Label4(0).Caption = "Insertando registros en temporal"
    DoEvents
    
    
    'insertamos los registros en la temporal
    Sql = "insert into tmpslhfac (codusu,letraser,numfactu,fecfactu,numlinea,numalbar,fecalbar,horalbar,numtarje,codartic,cantidad,preciove,implinea,precioinicial,codsocio) "
    Sql = Sql & " select " & vSesion.Codigo & ",letraser,numfactu,fecfactu,numlinea,numalbar,fecalbar,horalbar,numtarje,codartic,cantidad,preciove,implinea,precioinicial,0  from slhfac where " & cWhere

    Conn.Execute Sql
    
    'insertamos los albararanes BONIFICA solo de las facturas que hay
    Sql = "insert into tmpslhfac (codusu,letraser,numfactu,fecfactu,numlinea,numalbar,fecalbar,horalbar,numtarje,codartic,cantidad,preciove,implinea,precioinicial,codsocio) "
    Sql = Sql & " select " & vSesion.Codigo & ",letraser,numfactu,fecfactu,numlinea,numalbar,fecalbar,horalbar,numtarje,codartic,cantidad,preciove,implinea,precioinicial,0  from slhfac "
    Sql = Sql & " where numalbar = 'BONIFICA'  "
    Sql = Sql & " and month(slhfac.fecfactu) = " & DBSet(txtCodigo(16).Text, "N") & " and year(slhfac.fecfactu) = " & DBSet(txtCodigo(17).Text, "N")
    Sql = Sql & " and (letraser, numfactu, fecfactu) in (select letraser, numfactu, fecfactu from tmpslhfac where codusu = " & vSesion.Codigo & ")"
    
    Conn.Execute Sql
    
    Me.Label4(0).Caption = "Insertando socios en temporal"
    DoEvents
    
    ' socio de la factura
    Sql = "update tmpslhfac, schfac "
    Sql = Sql & " set tmpslhfac.codsocio = schfac.codsocio "
    Sql = Sql & " where tmpslhfac.letraser = schfac.letraser and tmpslhfac.numfactu= schfac.numfactu "
    Sql = Sql & " and tmpslhfac.fecfactu = schfac.fecfactu "
    Sql = Sql & " and tmpslhfac.codusu = " & DBSet(vSesion.Codigo, "N")
    
    Conn.Execute Sql
    
    
    Me.Label4(0).Caption = "Insertando Precios Plat en temporal"
    DoEvents
    
    
    'campo1 0= PL
    '       1= no plat
                                          'PL,    precio, socio,  litros,   BI        BI Coste
    Sql = "insert into tmpinformes (codusu,campo1,precio2,codigo1,importe1, importe2, importe3) "
    CadInsert = Sql
    
    Sql = Sql & " select " & vSesion.Codigo & ", 0, preciove, tmpslhfac.codsocio, sum(cantidad), round(sum(implinea) / (1 + (" & DBSet(PorcIva, "N") & "/ 100)), 2), round(sum(cantidad) * " & DBSet(PrCoste, "N") & ",2) "
    Sql = Sql & " from tmpslhfac inner join ssocio on  tmpslhfac.codsocio = ssocio.codsocio"
    Sql = Sql & " where precioinicial <> 0 "
    Sql = Sql & " and codusu = " & vSesion.Codigo
    Sql = Sql & " group by 1,2,3,4"
    
    Conn.Execute Sql
    
    
    'metemos los que no son precio plat
    
    '1º sin descuento a pie de factura
    Me.Pb1.visible = False
    Label4(0).Caption = "Procesando sin descuento a pie de factura"
    DoEvents
    
    Sql = "select * from tmpslhfac where precioinicial is null "
    Sql = Sql & " and codusu = " & vSesion.Codigo
    Sql = Sql & " and not (letraser, numfactu, fecfactu) in (select letraser, numfactu, fecfactu from tmpslhfac where codusu = " & vSesion.Codigo & " and numalbar = 'BONIFICA') "
    Sql = Sql & " order by codsocio, numalbar, codartic "
    
    nRegs = TotalRegistrosConsulta(Sql)
    
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    CargarProgres Me.Pb1, CInt(nRegs)
    Me.Pb1.visible = True
    
    
    If Not Rs.EOF Then
        AlbarAnt = Rs!numalbar
        PrecioAnt = Rs!preciove
        LitrosAnt = 0
        SocioAnt = Rs!codsocio
        ArticAnt = Rs!codArtic
        Importe = 0
    End If
    
    CadValues = ""
    While Not Rs.EOF
        
        Label4(0).Caption = "Procesando Albarán " & Format(Rs!numalbar, "0000000")
        IncrementarProgres Me.Pb1, 1
        DoEvents
           
        If AlbarAnt <> DBLet(Rs!numalbar) Or SocioAnt <> DBLet(Rs!codsocio) Then
        
            If ArticAnt <> ArtDto Then
                PrecioDto = 0
            Else
                PrecioDto = PrecioAnt
            End If
            
            ' calculamos el importe sin el iva
            Importe = Round2(Importe / (1 + (PorcIva / 100)), 2)
            
            ImpCoste = Round2(LitrosAnt * PrCoste, 2)
        
            Sql2 = "select count(*) from tmpinformes where codusu = " & DBSet(vSesion.Codigo, "N")
            Sql2 = Sql2 & " and campo1 = 1 and precio2 = " & DBSet(PrecioDto, "N")
            Sql2 = Sql2 & " and codigo1 = " & DBSet(SocioAnt, "N")
            
            If TotalRegistros(Sql2) = 0 Then
                ' insertamos
                Sql2 = CadInsert & " values (" & vSesion.Codigo & ",1," & DBSet(PrecioDto, "N") & "," & DBSet(SocioAnt, "N") & ","
                Sql2 = Sql2 & DBSet(LitrosAnt, "N") & "," & DBSet(Importe, "N") & "," & DBSet(ImpCoste, "N") & ")"
                
                Conn.Execute Sql2
            Else
                ' updateamos
                Sql2 = "update tmpinformes "
                Sql2 = Sql2 & " set importe1 = importe1 + " & DBSet(LitrosAnt, "N")
                Sql2 = Sql2 & ", importe2 = importe2 + " & DBSet(Importe, "N")
                Sql2 = Sql2 & ", importe3 = importe3 + " & DBSet(ImpCoste, "N")
                Sql2 = Sql2 & " where codusu = " & vSesion.Codigo
                Sql2 = Sql2 & " and campo1 = 1 and precio2 = " & DBSet(PrecioDto, "N")
                Sql2 = Sql2 & " and codigo1 = " & DBSet(SocioAnt, "N")
                
                Conn.Execute Sql2
            End If
                
            AlbarAnt = Rs!numalbar
            ArticAnt = Rs!codArtic
            PrecioAnt = Rs!preciove
            
            SocioAnt = Rs!codsocio
            
            LitrosAnt = 0
            Importe = 0
            ImpCoste = 0
            
        End If
        
        Importe = Importe + DBLet(Rs!ImpLinea)
        
        If DBLet(Rs!codArtic) <> ArtDto Then
            LitrosAnt = LitrosAnt + DBLet(Rs!cantidad)
        End If
        
        
'        SocioAnt = DBLet(Rs!codsocio)
        PrecioAnt = DBLet(Rs!preciove)
        ArticAnt = DBLet(Rs!codArtic)
        
        Rs.MoveNext
    Wend
    Set Rs = Nothing
    
    
    ' ultimo registro
    If nRegs <> 0 Then
        If ArticAnt <> ArtDto Then
            PrecioDto = 0
        Else
            PrecioDto = PrecioAnt
        End If
        
        ' calculamos el importe sin el iva
        Importe = Round2(Importe / (1 + (PorcIva / 100)), 2)
        
        ImpCoste = Round2(LitrosAnt * PrCoste, 2)
    
        Sql2 = "select count(*) from tmpinformes where codusu = " & DBSet(vSesion.Codigo, "N")
        Sql2 = Sql2 & " and campo1 = 1 and precio2 = " & DBSet(PrecioDto, "N")
        Sql2 = Sql2 & " and codigo1 = " & DBSet(SocioAnt, "N")
        
        If TotalRegistros(Sql2) = 0 Then
            ' insertamos
            Sql2 = CadInsert & " values (" & vSesion.Codigo & ",1," & DBSet(PrecioDto, "N") & "," & DBSet(SocioAnt, "N") & ","
            Sql2 = Sql2 & DBSet(LitrosAnt, "N") & "," & DBSet(Importe, "N") & "," & DBSet(ImpCoste, "N") & ")"
            
            Conn.Execute Sql2
        Else
            ' updateamos
            Sql2 = "update tmpinformes "
            Sql2 = Sql2 & " set importe1 = importe1 + " & DBSet(LitrosAnt, "N")
            Sql2 = Sql2 & ", importe2 = importe2 + " & DBSet(Importe, "N")
            Sql2 = Sql2 & ", importe3 = importe3 + " & DBSet(ImpCoste, "N")
            Sql2 = Sql2 & " where codusu = " & vSesion.Codigo
            Sql2 = Sql2 & " and campo1 = 1 and precio2 = " & DBSet(PrecioDto, "N")
            Sql2 = Sql2 & " and codigo1 = " & DBSet(SocioAnt, "N")
            
            Conn.Execute Sql2
        End If
    End If
    
    
'*************
    '2º CON descuento a pie de factura
    
    Me.Pb1.visible = False
    Label4(0).Caption = "Procesando con descuento a pie de factura"
    DoEvents
    
    Sql = "select * from tmpslhfac where precioinicial is null "
    Sql = Sql & " and codusu = " & vSesion.Codigo
    Sql = Sql & " and (letraser, numfactu, fecfactu) in (select letraser, numfactu, fecfactu from tmpslhfac where codusu = " & vSesion.Codigo & " and numalbar = 'BONIFICA') "
    Sql = Sql & " order by codsocio, numfactu, codartic "
    
    nRegs = TotalRegistrosConsulta(Sql)
    
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    CargarProgres Me.Pb1, CInt(nRegs)
    Me.Pb1.visible = True
    
    
    
    If Not Rs.EOF Then
        FactuAnt = Rs!numfactu
        PrecioAnt = Rs!preciove
        LitrosAnt = 0
        SocioAnt = Rs!codsocio
        ArticAnt = Rs!codArtic
        Importe = 0
    End If
    
    CadValues = ""
    While Not Rs.EOF
        
        Label4(0).Caption = "Procesando Albarán " & Format(Rs!numalbar, "0000000")
        IncrementarProgres Me.Pb1, 1
        DoEvents
        
'        If AlbarAnt <> DBLet(Rs!numalbar) And ArticAnt = ArtDto Then
'            PrecioDto = PrecioAnt
'
'            Sql22 = "select sum(implinea) from tmpslhfac where codusu = " & DBSet(vSesion.Codigo, "N")
'            Sql22 = Sql22 & " and numfactu = " & DBSet(FactuAnt, "N")
'            Sql22 = Sql22 & " and codsocio = " & DBSet(SocioAnt, "N")
'            Sql22 = Sql22 & " and numalbar = " & DBSet(AlbarAnt, "T")
'
'            Importe2 = DevuelveValor(Sql22)
'
''            LitrosAnt = LitrosAnt - Litros2
'
'            Importe = Importe - Importe2
'
'            ' calculamos el importe sin el iva
'            Importe2 = Round2(Importe2 / (1 + (PorcIva / 100)), 2)
'
'            ImpCoste = Round2(Litros2 * PrCoste, 2)
'
'            Sql2 = "select count(*) from tmpinformes where codusu = " & DBSet(vSesion.Codigo, "N")
'            Sql2 = Sql2 & " and campo1 = 1 and precio2 = " & DBSet(PrecioDto, "N")
'            Sql2 = Sql2 & " and codigo1 = " & DBSet(SocioAnt, "N")
'
'            If TotalRegistros(Sql2) = 0 Then
'                ' insertamos
'                Sql2 = CadInsert & " values (" & vSesion.Codigo & ",1," & DBSet(PrecioDto, "N") & "," & DBSet(SocioAnt, "N") & ","
'                Sql2 = Sql2 & DBSet(Litros2, "N") & "," & DBSet(Importe2, "N") & "," & DBSet(ImpCoste, "N") & ")"
'
'                Conn.Execute Sql2
'            Else
'                ' updateamos
'                Sql2 = "update tmpinformes "
'                Sql2 = Sql2 & " set importe1 = importe1 + " & DBSet(Litros2, "N")
'                Sql2 = Sql2 & ", importe2 = importe2 + " & DBSet(Importe2, "N")
'                Sql2 = Sql2 & ", importe3 = importe3 + " & DBSet(ImpCoste, "N")
'                Sql2 = Sql2 & " where codusu = " & vSesion.Codigo
'                Sql2 = Sql2 & " and campo1 = 1 and precio2 = " & DBSet(PrecioDto, "N")
'                Sql2 = Sql2 & " and codigo1 = " & DBSet(SocioAnt, "N")
'
'                Conn.Execute Sql2
'            End If
'
'            AlbarAnt = Rs!numalbar
'            FactuAnt = Rs!numfactu
'            ArticAnt = Rs!codartic
'            PrecioAnt = Rs!preciove
'
'            SocioAnt = Rs!codsocio
'
'            Litros2 = 0
'            Importe2 = 0
'            ImpCoste = 0
'
'
'        Else
        
            If FactuAnt <> DBLet(Rs!numfactu) Or SocioAnt <> DBLet(Rs!codsocio) Then
                If ArticAnt <> vParamAplic.ArticDto Then
                    PrecioDto = 0
                Else
                    PrecioDto = PrecioAnt * (-1)
                End If
                
                ' calculamos el importe sin el iva
                Importe = Round2(Importe / (1 + (PorcIva / 100)), 2)
                
                ImpCoste = Round2(LitrosAnt * PrCoste, 2)
            
                Sql2 = "select count(*) from tmpinformes where codusu = " & DBSet(vSesion.Codigo, "N")
                Sql2 = Sql2 & " and campo1 = 1 and precio2 = " & DBSet(PrecioDto, "N")
                Sql2 = Sql2 & " and codigo1 = " & DBSet(SocioAnt, "N")
                
                If TotalRegistros(Sql2) = 0 Then
                    ' insertamos
                    Sql2 = CadInsert & " values (" & vSesion.Codigo & ",1," & DBSet(PrecioDto, "N") & "," & DBSet(SocioAnt, "N") & ","
                    Sql2 = Sql2 & DBSet(LitrosAnt, "N") & "," & DBSet(Importe, "N") & "," & DBSet(ImpCoste, "N") & ")"
                    
                    Conn.Execute Sql2
                Else
                    ' updateamos
                    Sql2 = "update tmpinformes "
                    Sql2 = Sql2 & " set importe1 = importe1 + " & DBSet(LitrosAnt, "N")
                    Sql2 = Sql2 & ", importe2 = importe2 + " & DBSet(Importe, "N")
                    Sql2 = Sql2 & ", importe3 = importe3 + " & DBSet(ImpCoste, "N")
                    Sql2 = Sql2 & " where codusu = " & vSesion.Codigo
                    Sql2 = Sql2 & " and campo1 = 1 and precio2 = " & DBSet(PrecioDto, "N")
                    Sql2 = Sql2 & " and codigo1 = " & DBSet(SocioAnt, "N")
                    
                    Conn.Execute Sql2
                End If
                    
                AlbarAnt = Rs!numalbar
                FactuAnt = Rs!numfactu
                ArticAnt = Rs!codArtic
                PrecioAnt = Rs!preciove
                
                SocioAnt = Rs!codsocio
                
                LitrosAnt = 0
                Importe = 0
                ImpCoste = 0
                
            End If
            
'        End If
        
        Litros2 = DBLet(Rs!cantidad)
        
        If DBLet(Rs!codArtic) <> vParamAplic.ArticDto Then
            If DBLet(Rs!codArtic) <> ArtDto Then
                LitrosAnt = LitrosAnt + DBLet(Rs!cantidad)
            End If
            Importe = Importe + DBLet(Rs!ImpLinea)
        Else
            Importe = Importe + Round2(Rs!preciove * LitrosAnt, 2)
        End If
        
        
'        SocioAnt = DBLet(Rs!codsocio)
        PrecioAnt = DBLet(Rs!preciove)
        ArticAnt = DBLet(Rs!codArtic)
'        AlbarAnt = DBLet(Rs!numalbar)
        
        Rs.MoveNext
    Wend
    Set Rs = Nothing
    
    
    ' ultimo registro
    If nRegs <> 0 Then
'        If ArticAnt = ArtDto Then
'            PrecioDto = PrecioAnt
'
'            Sql22 = "select sum(implinea) from tmpslhfac where codusu = " & DBSet(vSesion.Codigo, "N")
'            Sql22 = Sql22 & " and numfactu = " & DBSet(FactuAnt, "N")
'            Sql22 = Sql22 & " and codsocio = " & DBSet(SocioAnt, "N")
'            Sql22 = Sql22 & " and numalbar = " & DBSet(AlbarAnt, "T")
'
'            Importe2 = DevuelveValor(Sql22)
'
'            LitrosAnt = LitrosAnt - Litros2
'
'            Importe = Importe - Importe2
'
'            ' calculamos el importe sin el iva
'            Importe2 = Round2(Importe2 / (1 + (PorcIva / 100)), 2)
'
'            ImpCoste = Round2(Litros2 * PrCoste, 2)
'
'            Sql2 = "select count(*) from tmpinformes where codusu = " & DBSet(vSesion.Codigo, "N")
'            Sql2 = Sql2 & " and campo1 = 1 and precio2 = " & DBSet(PrecioDto, "N")
'            Sql2 = Sql2 & " and codigo1 = " & DBSet(SocioAnt, "N")
'
'            If TotalRegistros(Sql2) = 0 Then
'                ' insertamos
'                Sql2 = CadInsert & " values (" & vSesion.Codigo & ",1," & DBSet(PrecioDto, "N") & "," & DBSet(SocioAnt, "N") & ","
'                Sql2 = Sql2 & DBSet(Litros2, "N") & "," & DBSet(Importe2, "N") & "," & DBSet(ImpCoste, "N") & ")"
'
'                Conn.Execute Sql2
'            Else
'                ' updateamos
'                Sql2 = "update tmpinformes "
'                Sql2 = Sql2 & " set importe1 = importe1 + " & DBSet(Litros2, "N")
'                Sql2 = Sql2 & ", importe2 = importe2 + " & DBSet(Importe2, "N")
'                Sql2 = Sql2 & ", importe3 = importe3 + " & DBSet(ImpCoste, "N")
'                Sql2 = Sql2 & " where codusu = " & vSesion.Codigo
'                Sql2 = Sql2 & " and campo1 = 1 and precio2 = " & DBSet(PrecioDto, "N")
'                Sql2 = Sql2 & " and codigo1 = " & DBSet(SocioAnt, "N")
'
'                Conn.Execute Sql2
'            End If
'
'            AlbarAnt = Rs!numalbar
'            FactuAnt = Rs!numfactu
'            ArticAnt = Rs!codartic
'            PrecioAnt = Rs!preciove
'
'            SocioAnt = Rs!codsocio
'
'            Litros2 = 0
'            Importe2 = 0
'            ImpCoste = 0
'
'
'        Else
'
            If ArticAnt <> vParamAplic.ArticDto Then
                PrecioDto = 0
            Else
                PrecioDto = PrecioAnt * (-1)
            End If
            
            ' calculamos el importe sin el iva
            Importe = Round2(Importe / (1 + (PorcIva / 100)), 2)
            
            ImpCoste = Round2(LitrosAnt * PrCoste, 2)
        
            Sql2 = "select count(*) from tmpinformes where codusu = " & DBSet(vSesion.Codigo, "N")
            Sql2 = Sql2 & " and campo1 = 1 and precio2 = " & DBSet(PrecioDto, "N")
            Sql2 = Sql2 & " and codigo1 = " & DBSet(SocioAnt, "N")
            
            If TotalRegistros(Sql2) = 0 Then
                ' insertamos
                Sql2 = CadInsert & " values (" & vSesion.Codigo & ",1," & DBSet(PrecioDto, "N") & "," & DBSet(SocioAnt, "N") & ","
                Sql2 = Sql2 & DBSet(LitrosAnt, "N") & "," & DBSet(Importe, "N") & "," & DBSet(ImpCoste, "N") & ")"
                
                Conn.Execute Sql2
            Else
                ' updateamos
                Sql2 = "update tmpinformes "
                Sql2 = Sql2 & " set importe1 = importe1 + " & DBSet(LitrosAnt, "N")
                Sql2 = Sql2 & ", importe2 = importe2 + " & DBSet(Importe, "N")
                Sql2 = Sql2 & ", importe3 = importe3 + " & DBSet(ImpCoste, "N")
                Sql2 = Sql2 & " where codusu = " & vSesion.Codigo
                Sql2 = Sql2 & " and campo1 = 1 and precio2 = " & DBSet(PrecioDto, "N")
                Sql2 = Sql2 & " and codigo1 = " & DBSet(SocioAnt, "N")
                
                Conn.Execute Sql2
            End If
'        End If
    End If
    
    
'*************
    
    
    ' Indicamos el orden de tmpinformes por cantidad
    Label4(0).Caption = "Ordenando tabla "
    Pb1.visible = False
    DoEvents
    
    Sql = "select precio2, sum(importe1) from tmpinformes where codusu = " & vSesion.Codigo & " group by 1 order by 2 "
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    I = 0
    While Not Rs.EOF
        I = I + 1
        
        Sql2 = "update tmpinformes set importe4 = " & DBSet(I, "N") & " where codusu = " & vSesion.Codigo
        Sql2 = Sql2 & " and precio2 = " & DBSet(Rs!precio2, "N")
        
        Conn.Execute Sql2
        
        Rs.MoveNext
    Wend
    
    Set Rs = Nothing
    
    
'    'borramos de la temporal para no dejar datos
'    Sql = "delete from tmpslhfac where codusu = " & vSesion.Codigo
'    Conn.Execute Sql
    
    
    CargarTemporal = True
    Exit Function
    
eCargarTemporal:
    MuestraError Err.Number, "Cargar Temporal", Err.Description
End Function

Private Function DatosOk() As Boolean
'Dim Datos As String
Dim b As Boolean
Dim Sql As String
Dim Mens As String


    b = True
    
'    If txtCodigo(14).Text = "" Then
'        MsgBox "Debe introducir un artídulo. Reintroduzca.", vbExclamation
'        PonerFoco txtCodigo(14)
'        b = False
'    Else
'        If Not EsArticuloCombustible(txtCodigo(14).Text) Then
'            MsgBox "El artículo debe de ser de la familia de Combustibles. Reintroduzca.", vbExclamation
'            PonerFoco txtCodigo(14)
'            b = False
'        End If
'    End If

    
    'comprobamos mes anyo
    If b Then
        If txtCodigo(16).Text = "" Or txtCodigo(17).Text = "" Then
            MsgBox "Debe introducir Mes/Año de cálculo."
            PonerFoco txtCodigo(16)
            b = False
        End If
        If b Then
            If CInt(txtCodigo(16).Text) > 12 Or CInt(txtCodigo(16).Text) < 1 Then
                MsgBox "El mes ha de estar entre 1 y 12. Revise.", vbExclamation
                PonerFoco txtCodigo(16)
                b = False
            End If
            
            If b And CInt(txtCodigo(17).Text) < 0 Then
                MsgBox "Año incorrecto. Revise.", vbExclamation
                PonerFoco txtCodigo(17)
                b = False
            End If
        End If
    End If
         
         
         
    DatosOk = b
End Function





Private Sub cmdCancel_Click(Index As Integer)
    Unload Me
End Sub


Private Sub Form_Activate()
Dim IndiceFoco As Integer
Dim Mes As Integer
Dim Anyo As Long

    If PrimeraVez Then
        PrimeraVez = False
        
        Mes = Month(Now)
        Anyo = Year(Now)
        If Mes = 1 Then
            Anyo = Anyo - 1
            Mes = 12
        Else
            Mes = Mes - 1
        End If
        
        txtCodigo(16).Text = Format(Mes, "00")
        txtCodigo(17).Text = Format(Anyo, "0000")

        PonerFoco txtCodigo(14)
    
    End If
    
    Screen.MousePointer = vbDefault
End Sub


Private Sub Form_Load()
Dim h As Integer, w As Integer


'    'Icono del formulario
'    Me.Icon = frmPpal.Icon

    PrimeraVez = True
    limpiar Me
    
    'Ocultar todos los Frames de Formulario
    FrameMovArtic.visible = False
    
    CommitConexion
    
    CargarIconos
    
    cadTitulo = ""
    cadNomRPT = ""
    
    ListadosAlmacen h, w
    
    
    'Esto se consigue poneinedo el cancel en el opcion k corresponda
    Me.cmdCancel(3).Cancel = True
    Me.Width = w + 70
    Me.Height = h + 350
End Sub

Private Sub Form_Unload(Cancel As Integer)
    NumCod = ""
End Sub


Private Sub frmMens_DatoSeleccionado(CadenaSeleccion As String)
Dim I As Integer
Dim J As Integer
Dim K As Integer
Dim Artic As String
Dim Sql As String
Dim Sql2 As String

    
    If CadenaSeleccion <> "" Then
        Sql = CadenaSeleccion & ","
    
        I = Len(Sql)
        J = 1
        Do
            K = InStr(J, Sql, ",")
            If K > 0 Then
                Artic = Mid(Sql, J, K - J)
                
                If NomArtic = "" Then
                    NomArtic = DevuelveValor("select artdto from sartic where codartic = " & DBSet(Artic, "N"))
                    Articulo = Artic
                End If
                
                Articulos = Articulos & Artic & ","
                
                Sql2 = DevuelveValor("select artdto from sartic where codartic = " & DBSet(Artic, "N"))
   
                If Sql2 <> 0 Then
                     Articulos = Articulos & DBSet(Sql2, "N") & ","
                    If ArtDto = "" Then
                        ArtDto = Sql2
                    Else
                        If ArtDto <> Sql2 Then
                            OK = False
                        End If
                    End If
                End If
                
                J = K + 1
            End If
        Loop Until J > I
        
        OK = True
        Articulos = "(" & Mid(Articulos, 1, Len(Articulos) - 1) & ")"
    End If
    
    
End Sub

Private Sub frmMtoArticulos_DatoSeleccionado(CadenaSeleccion As String)
'Form de Mantenimiento de Articulos
    txtCodigo(indCodigo).Text = RecuperaValor(CadenaSeleccion, 1)
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmMtoClientes_DatoSeleccionado(CadenaSeleccion As String)
'Form de Mantenimiento de Clientes
    If indCodigo > 0 Then
        txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000000")
        txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
    End If
End Sub


Private Sub imgBuscarG_Click(Index As Integer)
'Buscar general: cada index llama a una tabla
    Screen.MousePointer = vbHourglass
    
    Select Case Index
            
        Case 11 'cod. ARTICULO
            Select Case Index
                Case 11, 12: indCodigo = Index + 3
            End Select
            Set frmMtoArticulos = New frmManArtic
            frmMtoArticulos.DatosADevolverBusqueda = "0|1|" 'Abrimos en Modo Busqueda
            frmMtoArticulos.Show vbModal
            Set frmMtoArticulos = Nothing
            
    End Select
    PonerFoco txtCodigo(indCodigo)
    Screen.MousePointer = vbDefault
End Sub

Private Sub txtCodigo_GotFocus(Index As Integer)
    ConseguirFoco txtCodigo(Index), 3
End Sub


Private Sub txtCodigo_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    KEYdown KeyCode
End Sub

Private Sub txtCodigo_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub txtCodigo_LostFocus(Index As Integer)
Dim tabla As String
Dim codCampo As String, nomcampo As String
Dim TipCampo As String, Formato As String
Dim Titulo As String
Dim EsNomCod As Boolean 'Si es campo Cod-Descripcion llama a PonerNombreDeCod


    'Quitar espacios en blanco por los lados
    txtCodigo(Index).Text = Trim(txtCodigo(Index).Text)

    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub

    EsNomCod = False
    
    Select Case Index
        
        Case 14 'Cod. ARTICULO
            EsNomCod = True
            tabla = "sartic"
            codCampo = "codartic"
            nomcampo = "nomartic"
            TipCampo = "T"
            Titulo = "Artículo"
            txtCodigo(Index).Text = UCase(txtCodigo(Index).Text)
        
        Case 16, 17 'Mes, anyo
            PonerFormatoEntero txtCodigo(Index)
        
        Case 2 'Precio de coste
            If txtCodigo(Index).Text <> "" Then
                PonerFormatoDecimal txtCodigo(2), 7
            End If
    
    End Select
    
    If EsNomCod Then
        If TipCampo = "N" Then
            If PonerFormatoEntero(txtCodigo(Index)) Then
                
                
                txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), tabla, nomcampo, codCampo)
'                txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), Tabla, NomCampo, codCampo, Titulo, TipCampo)
            
                If txtCodigo(Index).Text <> "" Then txtCodigo(Index).Text = Format(txtCodigo(Index).Text, Formato)
            Else
                txtNombre(Index).Text = ""
            End If
        Else
            txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), tabla, nomcampo, codCampo)
'            txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), Tabla, NomCampo, codCampo, Titulo, TipCampo)
        End If
    End If
    
   
End Sub


Private Sub frmB_Selecionado(CadenaDevuelta As String)
'Formulario para Busqueda

    If CadenaDevuelta <> "" Then
        HaDevueltoDatos = True
        Screen.MousePointer = vbHourglass
        Select Case OpcionListado
            Case 7, 8 'Informe Traspasos Almacen
                txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaDevuelta, 1), "0000000")
                PonerFoco txtCodigo(indCodigo)
            Case 9, 12, 13, 14, 15, 16, 17 '9: Informe Movimiento Articulos
                                'Inventario Articulos
                                '14: Actualizar diferencias Stock Inventariado
                                '16: Listado Valoracion stock inventariado
                txtCodigo(indCodigo).Text = RecuperaValor(CadenaDevuelta, 1)
                txtNombre(indCodigo).Text = RecuperaValor(CadenaDevuelta, 2)
                PonerFoco txtCodigo(indCodigo)
        End Select
    End If
    Screen.MousePointer = vbDefault
End Sub




Private Function PonerFormulaYParametrosInf9() As Boolean
Dim Cad As String
Dim todosMarcados As Boolean
Dim devuelve As String
Dim I As Byte

    PonerFormulaYParametrosInf9 = False
    InicializarVbles
    
    'Parametro EMPRESA
    cadParam = "|pNomEmpre=""" & vEmpresa.nomEmpre & """|"
    numParam = 1
        
    'Cadena para seleccion Desde y Hasta ARTICULO
    If txtCodigo(2).Text <> "" Or txtCodigo(3).Text <> "" Then
        Codigo = "{slhfac.fecfactu}"
        devuelve = "pDHFecha=""Fecha Factura: "
        If Not PonerDesdeHasta(Codigo, "F", 2, 3, devuelve) Then Exit Function
    End If
        
    'Cadena para seleccion Desde y Hasta ARTICULO
    If txtCodigo(14).Text <> "" Or txtCodigo(15).Text <> "" Then
        Codigo = "{slhfac.codartic}"
        devuelve = "pDHArticulo=""Artículo: "
        If Not PonerDesdeHasta(Codigo, "N", 14, 15, devuelve) Then Exit Function
    End If
                    
    'Cadena para seleccion Desde y Hasta FAMILIA
    If txtCodigo(16).Text <> "" Or txtCodigo(17).Text <> "" Then
        Codigo = "{sartic.codfamia}"
        devuelve = "pDHFamilia=""Familia: "
        If Not PonerDesdeHasta(Codigo, "N", 16, 17, devuelve) Then Exit Function
    End If
    
    PonerFormulaYParametrosInf9 = True
    
End Function



Private Function AnyadirParametroDH(Cad As String, indD As Byte, indH As Byte) As String
On Error Resume Next
    
     If txtCodigo(indD).Text <> "" Then
        Cad = Cad & "desde " & txtCodigo(indD).Text
        If txtNombre(indD).Text <> "" Then Cad = Cad & " - " & txtNombre(indD).Text
    End If
    If txtCodigo(indH).Text <> "" Then
        Cad = Cad & "  hasta " & txtCodigo(indH).Text
        If txtNombre(indH).Text <> "" Then Cad = Cad & " - " & txtNombre(indH).Text
    End If
    
    AnyadirParametroDH = Cad
    If Err.Number <> 0 Then Err.Clear
End Function


Private Sub InicializarVbles()
    cadFormula = ""
    cadSelect = ""
    cadParam = ""
    numParam = 0
'    cadTitulo = ""
'    cadNomRPT = ""
    conSubRPT = False
End Sub


Private Function PonerDesdeHasta(campo As String, tipo As String, indD As Byte, indH As Byte, param As String) As Boolean
Dim devuelve As String
Dim Cad As String

    PonerDesdeHasta = False
    devuelve = CadenaDesdeHasta(txtCodigo(indD).Text, txtCodigo(indH).Text, campo, tipo)
    If devuelve = "Error" Then Exit Function
    If Not AnyadirAFormula(cadFormula, devuelve) Then Exit Function
    
    If tipo <> "F" Then
        'Fecha para Crystal Report
        If Not AnyadirAFormula(cadSelect, devuelve) Then Exit Function
    Else
        'Fecha para la Base de Datos
        Cad = CadenaDesdeHastaBD(txtCodigo(indD).Text, txtCodigo(indH).Text, campo, tipo)
        If Not AnyadirAFormula(cadSelect, Cad) Then Exit Function
    End If
    
    If devuelve <> "" Then
        If param <> "" Then
            'Parametro Desde/Hasta
            cadParam = cadParam & AnyadirParametroDH(param, indD, indH) & """|"
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
        .EnvioEMail = False
        .Opcion = OpcionListado
        .Titulo = cadTitulo
        .NombreRPT = cadNomRPT
        .ConSubInforme = conSubRPT
        .Show vbModal
    End With
End Sub


Private Sub AbrirFrmClientes()
'Clientes
    Set frmMtoClientes = New frmManClien
    frmMtoClientes.DatosADevolverBusqueda = "0|1|"
    frmMtoClientes.Show vbModal
    Set frmMtoClientes = Nothing
End Sub


Private Function ComprobarFechasConta(ind As Integer) As Boolean
'comprobar que el periodo de fechas a contabilizar esta dentro del
'periodo de fechas del ejercicio de la contabilidad
Dim FechaIni As String, FechaFin As String
Dim Cad As String
Dim Rs As ADODB.Recordset
    
On Error GoTo EComprobar

    ComprobarFechasConta = False
    
    If txtCodigo(ind).Text <> "" Then
        FechaIni = "Select fechaini,fechafin From parametros"
        Set Rs = New ADODB.Recordset
        Rs.Open FechaIni, ConnConta, adOpenForwardOnly, adLockPessimistic, adCmdText
    
        If Not Rs.EOF Then
            FechaIni = DBLet(Rs!FechaIni, "F")
            FechaFin = DateAdd("yyyy", 1, DBLet(Rs!FechaFin, "F"))
            'nos guardamos los valores
            Orden1 = FechaIni
            Orden2 = FechaFin
        
            If Not EntreFechas(FechaIni, txtCodigo(ind).Text, FechaFin) Then
                 Cad = "El período de contabilización debe estar dentro del ejercicio:" & vbCrLf & vbCrLf
                 Cad = Cad & "    Desde: " & FechaIni & vbCrLf
                 Cad = Cad & "    Hasta: " & FechaFin
                 MsgBox Cad, vbExclamation
                 txtCodigo(ind).Text = ""
            Else
                ComprobarFechasConta = True
            End If
        End If
        Rs.Close
        Set Rs = Nothing
    Else
        ComprobarFechasConta = True
    End If
    
EComprobar:
    If Err.Number <> 0 Then MuestraError Err.Number, "Comprobar Fechas", Err.Description
End Function




Private Sub ListadosAlmacen(h As Integer, w As Integer)
    'LISTADOS DE ALMACENES
    'Informe Movimiento Artículos
     w = 7995
     h = 4035
     PonerFrameVisible Me.FrameMovArtic, True, h, w
     indFrame = 3
     Codigo = "{smoval.codartic}"
     cadTitulo = "Margen Ventas por Cliente"
     conSubRPT = False
End Sub


Private Sub CargarIconos()
Dim I As Integer
    
    For I = 11 To 11
        Me.imgBuscarG(I).Picture = frmPpal.imgListImages16.ListImages(1).Picture
    Next I

End Sub
