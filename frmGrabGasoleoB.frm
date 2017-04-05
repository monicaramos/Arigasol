VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmGrabGasoleoB 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Grabación Fichero Gasóleo B - Modelo 544"
   ClientHeight    =   4035
   ClientLeft      =   45
   ClientTop       =   2430
   ClientWidth     =   7185
   Icon            =   "frmGrabGasoleoB.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4035
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
      Height          =   3795
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   6915
      Begin VB.Frame Frame1 
         Height          =   915
         Left            =   420
         TabIndex        =   20
         Top             =   1530
         Width           =   2745
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   2
            Left            =   1500
            MaxLength       =   10
            TabIndex        =   2
            Tag             =   "Código Postal|T|S|||clientes|codposta|||"
            Top             =   360
            Width           =   1050
         End
         Begin VB.Label Label4 
            Caption         =   "Entidad Gasol."
            ForeColor       =   &H00972E0B&
            Height          =   255
            Index           =   16
            Left            =   120
            TabIndex        =   21
            Top             =   360
            Width           =   1125
         End
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   3
         Left            =   1950
         MaxLength       =   10
         TabIndex        =   22
         Tag             =   "Código Postal|T|S|||clientes|codposta|||"
         Top             =   2130
         Width           =   1050
      End
      Begin VB.Frame FrameResultados2 
         Caption         =   "Resultados Llombai"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   1065
         Left            =   3360
         TabIndex        =   15
         Top             =   390
         Width           =   3315
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   7
            Left            =   180
            Locked          =   -1  'True
            MaxLength       =   10
            TabIndex        =   17
            Tag             =   "Código Postal|T|S|||clientes|codposta|||"
            Top             =   510
            Width           =   1200
         End
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   6
            Left            =   1620
            Locked          =   -1  'True
            MaxLength       =   10
            TabIndex        =   16
            Tag             =   "Código Postal|T|S|||clientes|codposta|||"
            Top             =   510
            Width           =   1500
         End
         Begin VB.Label Label4 
            Caption         =   "Nro.Beneficiarios"
            ForeColor       =   &H00972E0B&
            Height          =   255
            Index           =   5
            Left            =   180
            TabIndex        =   19
            Top             =   270
            Width           =   1215
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            Caption         =   "Importe Total"
            ForeColor       =   &H00972E0B&
            Height          =   255
            Index           =   4
            Left            =   2070
            TabIndex        =   18
            Top             =   270
            Width           =   1035
         End
      End
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   4470
         TabIndex        =   3
         Top             =   3180
         Width           =   975
      End
      Begin VB.Frame FrameResultados1 
         Caption         =   "Resultados Catadau"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   1065
         Left            =   3360
         TabIndex        =   9
         Top             =   1500
         Width           =   3315
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   5
            Left            =   1620
            Locked          =   -1  'True
            MaxLength       =   10
            TabIndex        =   12
            Tag             =   "Código Postal|T|S|||clientes|codposta|||"
            Top             =   510
            Width           =   1500
         End
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   4
            Left            =   180
            Locked          =   -1  'True
            MaxLength       =   10
            TabIndex        =   10
            Tag             =   "Código Postal|T|S|||clientes|codposta|||"
            Top             =   510
            Width           =   1200
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            Caption         =   "Importe Total"
            ForeColor       =   &H00972E0B&
            Height          =   255
            Index           =   3
            Left            =   2070
            TabIndex        =   13
            Top             =   270
            Width           =   1035
         End
         Begin VB.Label Label4 
            Caption         =   "Nro.Beneficiarios"
            ForeColor       =   &H00972E0B&
            Height          =   255
            Index           =   2
            Left            =   180
            TabIndex        =   11
            Top             =   270
            Width           =   1215
         End
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   0
         Left            =   1920
         MaxLength       =   10
         TabIndex        =   0
         Tag             =   "Código Postal|T|S|||clientes|codposta|||"
         Top             =   780
         Width           =   1050
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   1
         Left            =   1920
         MaxLength       =   10
         TabIndex        =   1
         Tag             =   "Código Postal|T|S|||clientes|codposta|||"
         Top             =   1140
         Width           =   1050
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   480
         Top             =   2730
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   5655
         TabIndex        =   4
         Top             =   3180
         Width           =   975
      End
      Begin MSComctlLib.ProgressBar Pb1 
         Height          =   255
         Left            =   450
         TabIndex        =   14
         Top             =   2700
         Width           =   6210
         _ExtentX        =   10954
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Label Label4 
         Caption         =   "Entidad Bancaria"
         ForeColor       =   &H00972E0B&
         Height          =   255
         Index           =   1
         Left            =   570
         TabIndex        =   23
         Top             =   2130
         Width           =   1305
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   1
         Left            =   1590
         Picture         =   "frmGrabGasoleoB.frx":000C
         ToolTipText     =   "Buscar fecha"
         Top             =   1140
         Width           =   240
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   0
         Left            =   1590
         Picture         =   "frmGrabGasoleoB.frx":0097
         ToolTipText     =   "Buscar fecha"
         Top             =   780
         Width           =   240
      End
      Begin VB.Label Label4 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   14
         Left            =   870
         TabIndex        =   8
         Top             =   1140
         Width           =   420
      End
      Begin VB.Label Label4 
         Caption         =   "Desde"
         Height          =   195
         Index           =   15
         Left            =   870
         TabIndex        =   7
         Top             =   780
         Width           =   465
      End
      Begin VB.Label Label4 
         Caption         =   "Fecha"
         ForeColor       =   &H00972E0B&
         Height          =   255
         Index           =   0
         Left            =   510
         TabIndex        =   6
         Top             =   510
         Width           =   1815
      End
   End
End
Attribute VB_Name = "frmGrabGasoleoB"
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
Private WithEvents frmCol As frmManCoope
Attribute frmCol.VB_VarHelpID = -1
Private WithEvents frmcli As frmManClien 'Clientes
Attribute frmcli.VB_VarHelpID = -1
Private WithEvents frmC As frmCal 'calendario fechas
Attribute frmC.VB_VarHelpID = -1
Private WithEvents frmMens As frmMensajes 'articulos de gasoleo B que van a declarar
Attribute frmMens.VB_VarHelpID = -1

Dim indCodigo As Integer 'indice para txtCodigo
Dim indFrame As Single 'nº de frame en el que estamos

'Se inicializan para cada Informe (tabla de BD a la que hace referencia
Dim tabla As String
Dim Codigo As String 'Código para FormulaSelection de Crystal Report
Dim TipCod As String
Dim Orden1 As String 'Campo de Ordenacion (por codigo) para Cristal Report
Dim Orden2 As String 'Campo de Ordenacion (por nombre) para Cristal Report

Dim PrimeraVez As Boolean

Dim Socios As Currency
Dim total As Currency
Dim Socios1 As Currency
Dim Total1 As Currency
Dim Socios2 As Currency
Dim Total2 As Currency

Dim CadenaArticulos As String

Dim Cadena As String


Private cadFormula As String 'Cadena con la FormulaSelection para Crystal Report
Private cadParam As String 'Cadena con los parametros para Crystal Report
Private numParam As Byte 'Numero de parametros que se pasan a Crystal Report
Private cadSelect As String 'Cadena para comprobar si hay datos antes de abrir Informe
Private cadTitulo As String 'Titulo para la ventana frmImprimir
Private cadNombreRPT As String 'Nombre del informe




Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, 0, cerrar
    If cerrar Then Unload Me
End Sub

Private Sub cmdAceptar_Click()
Dim I As Byte
Dim SQL As String
Dim NRegs As Long
Dim C As Object
Dim v_Cadena As String
Dim b As Boolean


    If DatosOk Then
        '[Monica]28/11/2011: Dejamos seleccionar que articulos de gasoleo B quieren declarar
        '                    solo seleccionamos los codigo de articulos si hay mas de un articulo de gasoleo B
        If TotalRegistros("select count(*) from sartic where tipogaso = 3") = 1 Then
            CadenaArticulos = "(1 = 1)"
        Else
            CadenaArticulos = ""
            
            Set frmMens = frmMensajes
            
            frmMens.OpcionMensaje = 15
            frmMens.cadWhere = "tipogaso = 3"
            frmMens.cmdEtiqEstan(1).Caption = "&Aceptar"
            frmMens.Show vbModal
            
            Set frmMens = Nothing
            
            If CadenaArticulos = "" Then CadenaArticulos = "sartic.codartic is null"
        End If
        'fin 28/11/2011
        
        
        '[Monica]21/01/2014: para el caso de alzira el codbanco es el de la tarjeta
        If vParamAplic.Cooperativa = 1 Then
            v_Cadena = "SELECT  cast(starje.codbanco as unsigned) codbanco, cast(starje.codsucur as unsigned) codsucur, starje.digcontr, starje.cuentaba, starje.iban, schfac.codsocio, sum(slhfac.implinea) "
            v_Cadena = v_Cadena & " from slhfac, schfac, starje where schfac.codsocio = starje.codsocio "
            If txtCodigo(0).Text <> "" Then v_Cadena = v_Cadena & " and slhfac.fecfactu >= " & DBSet(txtCodigo(0).Text, "F")
            If txtCodigo(1).Text <> "" Then v_Cadena = v_Cadena & " and slhfac.fecfactu <= " & DBSet(txtCodigo(1).Text, "F")
            v_Cadena = v_Cadena & " and slhfac.codartic in (select codartic from sartic where tipogaso = 3 and "
            v_Cadena = v_Cadena & CadenaArticulos & " ) "
            v_Cadena = v_Cadena & " and slhfac.numtarje = starje.numtarje and starje.tiptarje = 1 "
            v_Cadena = v_Cadena & " and slhfac.letraser = schfac.letraser and slhfac.numfactu = schfac.numfactu and slhfac.fecfactu = schfac.fecfactu "
            v_Cadena = v_Cadena & " group by 1, 2, 3, 4, 5, 6 "
            v_Cadena = v_Cadena & " order by 1, 2, 3, 4, 5, 6 "
            
            SQL = "select count(*) from (" & v_Cadena & ") as tabla "
        
        Else
            '[Monica]09/01/2013: Nueva cooperativa de Ribarroja
            If vParamAplic.Cooperativa = 3 Or vParamAplic.Cooperativa = 5 Then ' alzira o castellduc o Ribarroja
                v_Cadena = "SELECT  cast(ssocio.codbanco as unsigned) codbanco, cast(ssocio.codsucur as unsigned) codsucur, ssocio.digcontr, ssocio.cuentaba, ssocio.iban, schfac.codsocio, sum(slhfac.implinea) "
                v_Cadena = v_Cadena & " from slhfac, schfac, ssocio where schfac.codsocio = ssocio.codsocio "
                If txtCodigo(0).Text <> "" Then v_Cadena = v_Cadena & " and slhfac.fecfactu >= " & DBSet(txtCodigo(0).Text, "F")
                If txtCodigo(1).Text <> "" Then v_Cadena = v_Cadena & " and slhfac.fecfactu <= " & DBSet(txtCodigo(1).Text, "F")
                v_Cadena = v_Cadena & " and slhfac.codartic in (select codartic from sartic where tipogaso = 3 and "
    '[Monica]28/11/2011: añadida la CadenaArticulos linea siguiente
                v_Cadena = v_Cadena & CadenaArticulos & " ) "
    '[Monica]28/11/2011: añadida la concidicion de socios con tarjeta de gasoleo bonificado
                v_Cadena = v_Cadena & " and slhfac.numtarje in (select numtarje from starje where codsocio = schfac.codsocio and tiptarje = 1)"
                v_Cadena = v_Cadena & " and slhfac.letraser = schfac.letraser and slhfac.numfactu = schfac.numfactu and slhfac.fecfactu = schfac.fecfactu "
                v_Cadena = v_Cadena & " group by 1, 2, 3, 4, 5, 6 "
                v_Cadena = v_Cadena & " order by 1, 2, 6, 3, 4, 5 "
    
                
                SQL = "select count(*) from (" & v_Cadena & ") as tabla "
    
            Else                                ' regaixo
                
                v_Cadena = "SELECT  cast(ssocio.codbanco as unsigned) codbanco, cast(ssocio.codsucur as unsigned) codsucur, ssocio.digcontr, ssocio.cuentaba, ssocio.iban, schfacr.codsocio, sum(slhfacr.implinea) "
                v_Cadena = v_Cadena & " from slhfacr, schfacr, ssocio where schfacr.codsocio = ssocio.codsocio "
                If txtCodigo(0).Text <> "" Then v_Cadena = v_Cadena & " and slhfacr.fecfactu >= " & DBSet(txtCodigo(0).Text, "F")
                If txtCodigo(1).Text <> "" Then v_Cadena = v_Cadena & " and slhfacr.fecfactu <= " & DBSet(txtCodigo(1).Text, "F")
                v_Cadena = v_Cadena & " and slhfacr.codartic in (select codartic from sartic where tipogaso = 3 and "
    '[Monica]28/11/2011: añadida la CadenaArticulos linea siguiente
                v_Cadena = v_Cadena & CadenaArticulos & " ) "
    '[Monica]28/11/2011: añadida la concidicion de socios con tarjeta de gasoleo bonificado
                v_Cadena = v_Cadena & " and slhfacr.numtarje in (select numtarje from starje where codsocio = schfacr.codsocio and tiptarje = 1)"
                v_Cadena = v_Cadena & " and slhfacr.letraser = schfacr.letraser and slhfacr.numfactu = schfacr.numfactu and slhfacr.fecfactu = schfacr.fecfactu "
                v_Cadena = v_Cadena & " group by 1,2,3,4,5,6  "
                v_Cadena = v_Cadena & " order by 1,2,6,3,4,5 "
                
                SQL = "select count(*) from (" & v_Cadena & ") as tabla "
            
            End If
        End If
    
        NRegs = TotalRegistros(SQL)
    
        If NRegs <> 0 Then
'            '[Monica]21/03/2013: antes de generar el fichero he de ver si existen en entidaddom todos los bancos
'            If Not ComprobarBancos(v_Cadena) Then
'                MsgBox "No se ha realizado el proceso. Revise.", vbExclamation
''                MsgBox "No existen las Entidades Domiciliarias:" & vbCrLf & vbCrLf & CADENA, vbExclamation
'                Exit Sub
'            End If

            If Not ComprobarBancosNew(v_Cadena) Then
                MsgBox "No se ha realizado el proceso. Revise.", vbExclamation
                Exit Sub
            End If
            
            '[Monica]10/09/2015:comprobamos que el iban es correcto en las cuentas
            If Not ComprobarIBAN(v_Cadena) Then
                MsgBox "No se ha realizado el proceso. Revise.", vbExclamation
                Exit Sub
            End If
            
            
            Pb1.visible = True
            Pb1.Max = NRegs + 1
            Pb1.Value = 0
            
'            If vParamAplic.Cooperativa = 1 Then
'                b = GeneraFicheroNewAlz(v_Cadena)
'            Else
'                b = GeneraFicheroNew(v_Cadena)
'            End If


            '[Monica]19/01/2015: todos pasan a generar el mismo fichero de modelo 544 con las modificaciones de Junio de 2014
            b = GeneraFicheroNewAlz(v_Cadena)
            
            If b Then
                If CopiarFichero Then
                    MsgBox "Proceso realizado correctamente", vbExclamation
                   
                    NRegs = TotalRegistros("select count(*) from tmpinformes where codusu = " & vSesion.Codigo)
                    If NRegs <> 0 Then
                        cadParam = cadParam & "|pEmpresa=""" & vEmpresa.nomEmpre & """|"
                        numParam = numParam + 1
                        
                        cadFormula = "{tmpinformes.codusu} = " & vSesion.Codigo
                        
                        cadTitulo = "Resumen Fichero Modelo 544"
                        cadNombreRPT = "rResumen544.rpt"
                        
                        LlamarImprimir
                    End If
                   
                    Pb1.visible = False
                End If
            End If
        Else
            MsgBox "No hay registros para generar el fichero", vbExclamation
        End If
    End If
    
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
'     Me.imgBuscar(0).Picture = frmPpal.imgListImages16.ListImages(1).Picture
'     Me.imgBuscar(1).Picture = frmPpal.imgListImages16.ListImages(1).Picture
'     Me.imgBuscar(6).Picture = frmPpal.imgListImages16.ListImages(1).Picture

    '###Descomentar
'    CommitConexion

    FrameCobrosVisible True, h, w
    indFrame = 5
    tabla = "slhfac"
    FrameResultados1.visible = False
    FrameResultados2.visible = False
    Pb1.visible = False
'[Monica]26/03/2013:quito esto
'    '[Monica]09/01/2013: Nueva cooperativa de Ribarroja
'    Frame1.visible = (vParamAplic.Cooperativa = 1 Or vParamAplic.Cooperativa = 3 Or vParamAplic.Cooperativa = 5)
'    Frame1.Enabled = (vParamAplic.Cooperativa = 1 Or vParamAplic.Cooperativa = 3 Or vParamAplic.Cooperativa = 5)
    'Esto se consigue poneinedo el cancel en el opcion k corresponda
    Me.cmdCancel.Cancel = True
    Me.Width = w + 70
    Me.Height = h + 350
End Sub

Private Sub frmC_Selec(vFecha As Date)
 'Fecha
    txtCodigo(CByte(imgFec(0).Tag)).Text = Format(vFecha, "dd/MM/yyyy")
End Sub

Private Sub frmMens_DatoSeleccionado(CadenaSeleccion As String)

    If CadenaSeleccion <> "" Then
        CadenaArticulos = " sartic.codartic in (" & CadenaSeleccion & ")"
    Else
        CadenaArticulos = " sartic.codartic is null  "
    End If
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
    PonerFoco txtCodigo(CByte(imgFec(0).Tag) + 2)
    ' ***************************
End Sub

Private Sub ListView1_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub Optcodigo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then 'ENTER
    ElseIf KeyAscii = 27 Then Unload Me  'ESC
    End If
End Sub

Private Sub OptNombre_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then 'ENTER
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
            Case 0: KEYFecha KeyAscii, 0 'FECHA desde
            Case 1: KEYFecha KeyAscii, 1 'FECHA hasta
        End Select
    Else
        KEYpress KeyAscii
    End If

End Sub

Private Sub KEYFecha(KeyAscii As Integer, indice As Integer)
    KeyAscii = 0
    imgFec_Click (indice)
End Sub

Private Sub txtCodigo_LostFocus(Index As Integer)
Dim Cad As String, cadTipo As String 'tipo cliente

    'Quitar espacios en blanco por los lados
    txtCodigo(Index).Text = Trim(txtCodigo(Index).Text)

    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub

    Select Case Index
        Case 0, 1 'FECHA
            If txtCodigo(Index).Text <> "" Then PonerFormatoFecha txtCodigo(Index)

        Case 2, 3 ' ENTIDADES
            If txtCodigo(Index).Text <> "" Then txtCodigo(Index).Text = Format(txtCodigo(Index).Text, "0000")
    End Select
End Sub

Private Sub FrameCobrosVisible(visible As Boolean, ByRef h As Integer, ByRef w As Integer)
    Me.FrameCobros.visible = visible
    If visible = True Then
        Me.FrameCobros.Top = -90
        Me.FrameCobros.Left = 0
'        Me.FrameCobros.Height = 6015
'        Me.FrameCobros.Width = 6555
        w = Me.FrameCobros.Width
        h = Me.FrameCobros.Height
    End If
End Sub

Private Function GeneraFichero() As Boolean
Dim NFich As Integer
Dim Rs As ADODB.Recordset
Dim Cad As String
Dim SQL As String
Dim I As Integer
Dim vsocio As CSocio
Dim v_total As Currency
Dim v_total1 As Currency
Dim v_total2 As Currency
Dim v_lineas As Currency
Dim v_socios As Currency
Dim v_socios1 As Currency
Dim v_socios2 As Currency
Dim v_dombanco As String
Dim v_pobbanco As String
Dim AntCoope As Integer
Dim ActCoope As Integer
Dim Banco As Currency

    On Error GoTo EGen
    GeneraFichero = False

    NFich = FreeFile
    Open App.path & "\gasoleob.txt" For Output As #NFich

    Set Rs = New ADODB.Recordset
    
'    'partimos de la tabla de historico de facturas
'    sql = "SELECT  schfac.codcoope, schfac.codsocio, sum(slhfac.implinea) "
'    sql = sql & " from slhfac, schfac where 1 = 1 "
'    If txtCodigo(0).Text <> "" Then sql = sql & " and slhfac.fecfactu >= " & DBSet(txtCodigo(0).Text, "F")
'    If txtCodigo(1).Text <> "" Then sql = sql & " and slhfac.fecfactu <= " & DBSet(txtCodigo(1).Text, "F")
'    sql = sql & " and slhfac.codartic in (select codartic from sartic where tipogaso = 3) "
'
'    ' solo para el regaixo hay que añadir que el socio solo sea de la cooperativa 1 y 2
'    ' es decir de facturacion ajena
'    If vParamAplic.Cooperativa = 2 Then
'        sql = sql & " and schfac.codsocio in "
'        sql = sql & " (select codsocio from ssocio, scoope "
'        sql = sql & " where ssocio.codcoope = scoope.codcoope and scoope.tipfactu = 2)"
'    End If
    
    '[Monica]09/01/2013: Nueva cooperativa de Ribarroja
    If vParamAplic.Cooperativa = 1 Or vParamAplic.Cooperativa = 3 Or vParamAplic.Cooperativa = 5 Then ' alzira o castelduc o Ribarroja
        SQL = "SELECT  schfac.codcoope, schfac.codsocio, sum(slhfac.implinea) "
        SQL = SQL & " from slhfac, schfac where 1 = 1 "
        If txtCodigo(0).Text <> "" Then SQL = SQL & " and slhfac.fecfactu >= " & DBSet(txtCodigo(0).Text, "F")
        If txtCodigo(1).Text <> "" Then SQL = SQL & " and slhfac.fecfactu <= " & DBSet(txtCodigo(1).Text, "F")
        SQL = SQL & " and slhfac.codartic in (select codartic from sartic where tipogaso = 3 and "
        '[Monica]28/11/2011: añadida condicion
        SQL = SQL & CadenaArticulos & ") "
        '[Monica]28/11/2011: añadida la concidicion de socios con tarjeta de gasoleo bonificado
        SQL = SQL & " and slhfac.numtarje in (select numtarje from starje where codsocio = schfac.codsocio and tiptarje = 1)"
        SQL = SQL & " and slhfac.letraser = schfac.letraser and slhfac.numfactu = schfac.numfactu and slhfac.fecfactu = schfac.fecfactu "
        SQL = SQL & " group by 1, 2 "
        SQL = SQL & " order by 2"
    
    Else                                ' regaixo
        SQL = "SELECT  schfacr.codcoope, schfacr.codsocio, sum(slhfacr.implinea) "
        SQL = SQL & " from slhfacr, schfacr where 1 = 1 "
        If txtCodigo(0).Text <> "" Then SQL = SQL & " and slhfacr.fecfactu >= " & DBSet(txtCodigo(0).Text, "F")
        If txtCodigo(1).Text <> "" Then SQL = SQL & " and slhfacr.fecfactu <= " & DBSet(txtCodigo(1).Text, "F")
        SQL = SQL & " and slhfacr.codartic in (select codartic from sartic where tipogaso = 3 and "
        '[Monica]28/11/2011: añadida condicion
        SQL = SQL & CadenaArticulos & ") "
        '[Monica]28/11/2011: añadida la concidicion de socios con tarjeta de gasoleo bonificado
        SQL = SQL & " and slhfacr.numtarje in (select numtarje from starje where codsocio = schfacr.codsocio and tiptarje = 1) "
        SQL = SQL & " and slhfacr.letraser = schfacr.letraser and slhfacr.numfactu = schfacr.numfactu and slhfacr.fecfactu = schfacr.fecfactu "
        SQL = SQL & " group by 1,2 "
        SQL = SQL & " order by 1,2"
    
    End If
    
    
    Rs.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    
    v_lineas = 0
    
    '***************REGISTRO R1
    v_lineas = v_lineas + 1
    
    Cad = "E1"
    Cad = Cad & RellenaABlancos(Format(txtCodigo(1).Text, "yymmdd"), True, 12)
    '[Monica]09/01/2013: Nueva cooperativa de Ribarroja
    If vParamAplic.Cooperativa = 1 Or vParamAplic.Cooperativa = 3 Or vParamAplic.Cooperativa = 5 Then ' alzira o castelduc o Ribarroja
        Cad = Cad & RellenaABlancos(Format(txtCodigo(2).Text, "0000"), True, 144)
        Cad = Cad & Format(v_lineas, "0000000")
    Else ' regaixo
        Cad = Cad & RellenaABlancos("7056", True, 144)
        Cad = Cad & Format(v_lineas, "0000000")
    End If
    
    Print #NFich, Cad
    
    
    '***************REGISTRO R2
    v_lineas = v_lineas + 1
    
    Cad = RellenaABlancos("E2", True, 14)
    '[Monica]09/01/2013: Nueva cooperativa de Ribarroja
    If vParamAplic.Cooperativa = 1 Or vParamAplic.Cooperativa = 3 Or vParamAplic.Cooperativa = 5 Then
        Cad = Cad & Format(txtCodigo(2).Text, "0000")
        Cad = Cad & RellenaABlancos(Format(txtCodigo(3).Text, "0000"), True, 140)
    Else
        Set vsocio = New CSocio
        If vsocio.LeerDatos(DBLet(Rs.Fields(1).Value, "N")) Then
            If Rs.Fields(0).Value = 1 Then
                 v_dombanco = "MAESTRO TORRES RUIZ, 23"
                 v_pobbanco = "46195 LLOMBAI"
            Else
                 v_dombanco = "PLAZA ESPAÑA, 1"
                 v_pobbanco = "46196 CATADAU"
            End If
            Cad = Cad & "7056"
            Cad = Cad & RellenaABlancos(Format(vsocio.Banco, "0000"), True, 140)
            Set vsocio = Nothing
        End If
    End If
    Cad = Cad & Format(v_lineas, "0000000")
    Print #NFich, Cad
    
    AntCoope = Rs.Fields(0).Value
    ActCoope = AntCoope
    
    v_total = 0
    
    While Not Rs.EOF
    
        Set vsocio = New CSocio
        If vsocio.LeerDatos(DBLet(Rs.Fields(1).Value, "N")) Then
            ActCoope = Rs.Fields(0).Value
            
            If ActCoope <> AntCoope Then
                '[Monica]09/01/2013: Nueva cooperativa de Ribarroja
                If vParamAplic.Cooperativa = 1 Or vParamAplic.Cooperativa = 3 Or vParamAplic.Cooperativa = 5 Then
                    ' no hacemos nada en alzira
                Else
                    ' en Regaixo
                    'grabamos el E8 antes que el E2
                    v_lineas = v_lineas + 1
                    
                    Cad = RellenaABlancos("E8", True, 14)
                    Cad = Cad & RellenaABlancos(Format(Banco, "0000"), True, 20)
'                    cad = cad & RellenaABlancos(Format(vSocio.Banco, "0000"), True, 20)
                    If AntCoope = 1 Then
                        Cad = Cad & Format(Round2(v_total1 * 100, 0), "000000000000")
                        Cad = Cad & Format(v_socios1, "0000000000")
                        Cad = Cad & RellenaABlancos(Format(v_socios1 + 2, "0000000000"), True, 102)
                    Else
                        Cad = Cad & Format(Round2(v_total2 * 100, 0), "000000000000")
                        Cad = Cad & Format(v_socios2, "0000000000")
                        Cad = Cad & RellenaABlancos(Format(v_socios2 + 2, "0000000000"), True, 102)
                    End If
                    Cad = Cad & Format(v_lineas, "0000000")
                    Print #NFich, Cad
                    
                    v_lineas = v_lineas + 1
                    
                    Cad = RellenaABlancos("E2", True, 14)
                    If Rs.Fields(0).Value = 1 Then
                         v_dombanco = "MAESTRO TORRES RUIZ, 23"
                         v_pobbanco = "46195 LLOMBAI"
                    Else
                         v_dombanco = "PLAZA ESPAÑA, 1"
                         v_pobbanco = "46196 CATADAU"
                    End If
                    Cad = Cad & "7056"
                    Cad = Cad & RellenaABlancos(Format(vsocio.Banco, "0000"), True, 140)
                    Cad = Cad & Format(v_lineas, "0000000")
                    Print #NFich, Cad
                End If
                AntCoope = ActCoope
            End If
            
            Pb1.Value = Pb1.Value + 1
            
            v_lineas = v_lineas + 1
            
            v_socios = v_socios + 1
            If Rs.Fields(0).Value = 1 Then
                v_socios1 = v_socios1 + 1
            Else
                v_socios2 = v_socios2 + 1
            End If
            
            Cad = "E6"
            Cad = Cad & RellenaABlancos(vsocio.NIF, True, 10)
            Cad = Cad & "T "
            '[Monica]09/01/2013: Nueva cooperativa de Ribarroja
            If vParamAplic.Cooperativa = 1 Or vParamAplic.Cooperativa = 3 Or vParamAplic.Cooperativa = 5 Then
                Cad = Cad & Format(txtCodigo(3).Text, "0000")
            Else
                Cad = Cad & Format(vsocio.Banco, "0000")
            End If
            Cad = Cad & Format(vsocio.Sucursal, "0000")
            Cad = Cad & RellenaABlancos(vsocio.CuentaBan, True, 10)
            Cad = Cad & "46"
            Cad = Cad & Format(Round2(Rs.Fields(2).Value * 100, 0), "000000000000")
            Cad = Cad & RellenaABlancos(vsocio.Nombre, True, 36)
            '[Monica]09/01/2013: Nueva cooperativa de Ribarroja
            If vParamAplic.Cooperativa = 1 Or vParamAplic.Cooperativa = 3 Or vParamAplic.Cooperativa = 5 Then
                Cad = Cad & RellenaABlancos(vsocio.Domicilio, True, 36)
                Cad = Cad & RellenaABlancos(Trim(vsocio.CPostal) & " " & vsocio.POBLACION, True, 36)
            Else
                Cad = Cad & RellenaABlancos(v_dombanco, True, 36)
                Cad = Cad & RellenaABlancos(v_pobbanco, True, 36)
            End If
            Cad = Cad & RellenaABlancos(vsocio.Digcontrol, True, 4)
            Cad = Cad & Format(v_lineas, "0000000")
            
            Print #NFich, Cad
            
            v_total = v_total + Rs.Fields(2).Value
            If Rs.Fields(0).Value = 1 Then
                v_total1 = v_total1 + Rs.Fields(2).Value
            Else
                v_total2 = v_total2 + Rs.Fields(2).Value
            End If
            Banco = vsocio.Banco
            Set vsocio = Nothing
            
        End If
        
        Rs.MoveNext
    Wend
       
    '[Monica]09/01/2013: Nueva cooperativa de Ribarroja
    If vParamAplic.Cooperativa = 1 Or vParamAplic.Cooperativa = 3 Or vParamAplic.Cooperativa = 5 Then
        v_lineas = v_lineas + 1
        
        Cad = RellenaABlancos("E8", True, 14)
        Cad = Cad & RellenaABlancos(Format(txtCodigo(3).Text, "0000"), True, 20)
        Cad = Cad & Format(Round2(v_total * 100, 0), "000000000000")
        Cad = Cad & Format(v_socios, "0000000000")
        Cad = Cad & RellenaABlancos(Format(v_socios + 2, "0000000000"), True, 102)
        Cad = Cad & Format(v_lineas, "0000000")
        
        Print #NFich, Cad
    Else
        v_lineas = v_lineas + 1
        
        Cad = RellenaABlancos("E8", True, 14)
        Cad = Cad & RellenaABlancos(Format(Banco, "0000"), True, 20)
        If ActCoope = 1 Then
            Cad = Cad & Format(Round2(v_total1 * 100, 0), "000000000000")
            Cad = Cad & Format(v_socios1, "0000000000")
            Cad = Cad & RellenaABlancos(Format(v_socios1 + 2, "0000000000"), True, 102)
        Else
            Cad = Cad & Format(Round2(v_total2 * 100, 0), "000000000000")
            Cad = Cad & Format(v_socios2, "0000000000")
            Cad = Cad & RellenaABlancos(Format(v_socios2 + 2, "0000000000"), True, 102)
        End If
        Cad = Cad & Format(v_lineas, "0000000")
        
        Print #NFich, Cad
    End If
    
'    v_total = 0
    
    ' cargamos las variables a mostrar en el frame
    '[Monica]09/01/2013: Nueva cooperativa de Ribarroja
    If vParamAplic.Cooperativa = 1 Or vParamAplic.Cooperativa = 3 Or vParamAplic.Cooperativa = 5 Then
        Socios = v_socios
        total = v_total
    Else
        Socios1 = v_socios1
        Total1 = v_total1
        Socios2 = v_socios2
        Total2 = v_total2
        
    End If
    
    v_lineas = v_lineas + 1
    
    Cad = RellenaABlancos("E9", True, 34)
    '[Monica]09/01/2013: Nueva cooperativa de Ribarroja
    If vParamAplic.Cooperativa = 1 Or vParamAplic.Cooperativa = 5 Then
        Cad = Cad & Format(Round2(0 * 100, 0), "000000000000")
    Else
        Cad = Cad & Format(Round2(v_total * 100, 0), "000000000000")
    End If
    Cad = Cad & Format(v_socios, "0000000000")
    Cad = Cad & Format(v_lineas, "0000000000")
    
    '[Monica]09/01/2013: Nueva cooperativa de Ribarroja
    If vParamAplic.Cooperativa = 1 Or vParamAplic.Cooperativa = 3 Or vParamAplic.Cooperativa = 5 Then
        Cad = Cad & RellenaABlancos("001", True, 92)
    Else
        Cad = Cad & RellenaABlancos("002", True, 92)
    End If
    Cad = Cad & Format(v_lineas, "0000000")
    
    Print #NFich, Cad
       
    Rs.Close
    Set Rs = Nothing
    
    Close (NFich)
    If v_socios > 0 Then GeneraFichero = True
    Exit Function
EGen:
    Set Rs = Nothing
    Close (NFich)
    MuestraError Err.Number, Err.Description

End Function


Public Function CopiarFichero() As Boolean
Dim nomFich As String
Dim Cadena As String
On Error GoTo ecopiarfichero

    CopiarFichero = False
    ' abrimos el commondialog para indicar donde guardarlo
    Me.CommonDialog1.InitDir = App.path
    Me.CommonDialog1.DefaultExt = "txt"
    'cadena = Format(CDate(txtCodigo(2).Text), FormatoFecha)
    CommonDialog1.Filter = "Archivos txt|txt|"
    CommonDialog1.FilterIndex = 1
    
    CommonDialog1.FileName = "gasoleo.txt"
    
    Me.CommonDialog1.ShowSave
    If CommonDialog1.FileName <> "" Then
        FileCopy App.path & "\gasoleob.txt", CommonDialog1.FileName
        CopiarFichero = True
    End If

ecopiarfichero:
    If Err.Number <> 0 Then
        MuestraError Err.Number, Err.Description
    End If
    Err.Clear

End Function

Private Function RellenaABlancos(Cadena As String, PorLaDerecha As Boolean, longitud As Integer) As String
Dim Cad As String
    
    Cad = Space(longitud)
    If PorLaDerecha Then
        Cad = Cadena & Cad
        RellenaABlancos = Left(Cad, longitud)
    Else
        Cad = Cad & Cadena
        RellenaABlancos = Right(Cad, longitud)
    End If
    
End Function

Private Function DatosOk() As Boolean

    DatosOk = True
    
'[Monica]26/03/2013: quitamos esto
'    ' solo si es Alzira obligamos a introducir un valor en los campos entidad gasolinera y entidad bancaria
'    '[Monica]09/01/2013: Nueva cooperativa de Ribarroja
'    If vParamAplic.Cooperativa <> 1 And vParamAplic.Cooperativa <> 5 Then Exit Function
    
    If txtCodigo(2).Text = "" Then
        MsgBox "Debe introducir un valor en el campo entidad gasolinera.", vbExclamation
        PonerFoco txtCodigo(2)
        DatosOk = False
        Exit Function
    End If
    
'[Monica]26/03/2013: quitamos esto
'    If txtCodigo(3).Text = "" Then
'        MsgBox "Debe introducir un valor en el campo entidad bancaria.", vbExclamation
'        PonerFoco txtCodigo(3)
'        DatosOk = False
'        Exit Function
'    End If
    
End Function


Private Function GeneraFicheroNew(vCadselect As String) As Boolean
Dim NFich As Integer
Dim Rs As ADODB.Recordset
Dim RsBanco As ADODB.Recordset
Dim Cad As String
Dim SQL As String
Dim I As Integer
Dim vsocio As CSocio
Dim v_total As Currency
Dim v_total1 As Currency
Dim v_total2 As Currency
Dim v_lineas As Currency
Dim v_socios As Currency
Dim v_socios1 As Currency
Dim v_socios2 As Currency
Dim v_dombanco As String
Dim v_pobbanco As String
Dim AntBanco As Integer
Dim ActBanco As Integer
Dim Banco As Currency
Dim sql2 As String

Dim v_Entidades As Integer

Dim v_Sucursal As Integer
Dim v_Domicilio As String
Dim v_CPostal As String
Dim v_Poblacion As String
Dim v_Provincia As Integer

Dim IBAN As String

    On Error GoTo EGen
    GeneraFicheroNew = False

    SQL = "delete from tmpinformes where codusu = " & vSesion.Codigo
    Conn.Execute SQL


    NFich = FreeFile
    Open App.path & "\gasoleob.txt" For Output As #NFich

    Set Rs = New ADODB.Recordset
    
    SQL = vCadselect
    
    
'    '[Monica]09/01/2013: Nueva cooperativa de Ribarroja
'    If vParamAplic.Cooperativa = 1 Or vParamAplic.Cooperativa = 3 Or vParamAplic.Cooperativa = 5 Then ' alzira o castelduc o Ribarroja
'        Sql = "SELECT  ssocio.codbanco, schfac.codsocio, sum(slhfac.implinea) "
'        Sql = Sql & " from slhfac, schfac, ssocio where 1 = 1 "
'        If txtCodigo(0).Text <> "" Then Sql = Sql & " and slhfac.fecfactu >= " & DBSet(txtCodigo(0).Text, "F")
'        If txtCodigo(1).Text <> "" Then Sql = Sql & " and slhfac.fecfactu <= " & DBSet(txtCodigo(1).Text, "F")
'        Sql = Sql & " and slhfac.codartic in (select codartic from sartic where tipogaso = 3 and "
'        '[Monica]28/11/2011: añadida condicion
'        Sql = Sql & CadenaArticulos & ") "
'        '[Monica]28/11/2011: añadida la concidicion de socios con tarjeta de gasoleo bonificado
'        Sql = Sql & " and slhfac.numtarje in (select numtarje from starje where codsocio = schfac.codsocio and tiptarje = 1)"
'        Sql = Sql & " and slhfac.letraser = schfac.letraser and slhfac.numfactu = schfac.numfactu and slhfac.fecfactu = schfac.fecfactu "
'        Sql = Sql & " and schfac.codsocio = ssocio.codsocio "
'        Sql = Sql & " group by 1, 2 "
'        Sql = Sql & " order by 1, 2"
'    Else                                ' regaixo
'        Sql = "SELECT  ssocio.codbanco, schfacr.codsocio, sum(slhfacr.implinea) "
'        Sql = Sql & " from slhfacr, schfacr where 1 = 1 "
'        If txtCodigo(0).Text <> "" Then Sql = Sql & " and slhfacr.fecfactu >= " & DBSet(txtCodigo(0).Text, "F")
'        If txtCodigo(1).Text <> "" Then Sql = Sql & " and slhfacr.fecfactu <= " & DBSet(txtCodigo(1).Text, "F")
'        Sql = Sql & " and slhfacr.codartic in (select codartic from sartic where tipogaso = 3 and "
'        '[Monica]28/11/2011: añadida condicion
'        Sql = Sql & CadenaArticulos & ") "
'        '[Monica]28/11/2011: añadida la concidicion de socios con tarjeta de gasoleo bonificado
'        Sql = Sql & " and slhfacr.numtarje in (select numtarje from starje where codsocio = schfacr.codsocio and tiptarje = 1) "
'        Sql = Sql & " and slhfacr.letraser = schfacr.letraser and slhfacr.numfactu = schfacr.numfactu and slhfacr.fecfactu = schfacr.fecfactu "
'        Sql = Sql & " and schfacr.codsocio = ssocio.codsocio "
'        Sql = Sql & " group by 1,2 "
'        Sql = Sql & " order by 1,2"
'    End If
    
    Rs.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    
    v_lineas = 0
    
    '***************REGISTRO R1
    v_lineas = v_lineas + 1
    
    Cad = "E1"
    Cad = Cad & RellenaABlancos(Format(txtCodigo(1).Text, "yymmdd"), True, 12)
    Cad = Cad & RellenaABlancos(Format(txtCodigo(2).Text, "0000"), True, 144)
    Cad = Cad & Format(v_lineas, "0000000")
    
    Print #NFich, Cad
    
    AntBanco = Rs.Fields(0).Value
    ActBanco = Rs.Fields(0).Value
    
    v_Entidades = 0
    
    v_total = 0  'importe total
    v_total1 = 0 'importe por banco
    v_socios = 0 ' socios total
    v_socios1 = 0 'socios por banco
    
    While Not Rs.EOF
        ActBanco = DBLet(Rs.Fields(0).Value, "N")
    
        If ActBanco <> AntBanco Or v_socios1 = 0 Then
            If v_socios1 <> 0 Then
                '***************REGISTRO R8
                v_lineas = v_lineas + 1
                
                Cad = RellenaABlancos("E8", True, 14)
                Cad = Cad & RellenaABlancos(Format(AntBanco, "0000"), True, 20)
                Cad = Cad & Format(Round2(v_total1 * 100, 0), "000000000000")
                Cad = Cad & Format(v_socios1, "0000000000")
                Cad = Cad & RellenaABlancos(Format(v_socios1 + 2, "0000000000"), True, 102)
                Cad = Cad & Format(v_lineas, "0000000")
                
                Print #NFich, Cad
                
                sql2 = "insert into tmpinformes (codusu, codigo1, importe1, importe2) values "
                sql2 = sql2 & "(" & vSesion.Codigo & "," & DBSet(AntBanco, "N") & "," & DBSet(v_socios1, "N")
                sql2 = sql2 & "," & DBSet(v_total1, "N") & ")"
                
                Conn.Execute sql2
                
                v_total1 = 0
                v_socios1 = 0
                
                v_Entidades = v_Entidades + 1
                
            End If
        
            '***************REGISTRO R2
            v_lineas = v_lineas + 1
            
            'Nro total de socios del banco
            v_total1 = 0
            
            Cad = RellenaABlancos("E2", True, 14)
            Cad = Cad & Format(txtCodigo(2).Text, "0000")
            
            Cad = Cad & RellenaABlancos(Format(ActBanco, "0000"), True, 140)
            Cad = Cad & Format(v_lineas, "0000000")
            Print #NFich, Cad
            
            AntBanco = ActBanco
        
        End If
    
        Set vsocio = New CSocio
        If vsocio.LeerDatos(DBLet(Rs.Fields(2).Value, "N")) Then
            ' datos del banco
            v_Sucursal = 0
            v_Domicilio = ""
            v_CPostal = ""
            v_Poblacion = ""
            v_Provincia = 0
            
            v_Sucursal = vsocio.Sucursal
            v_Domicilio = vsocio.Domicilio
            v_CPostal = vsocio.CPostal
            v_Poblacion = vsocio.POBLACION
            v_Provincia = Mid(vsocio.CPostal, 1, 2)
            
            Pb1.Value = Pb1.Value + 1
            DoEvents
            v_lineas = v_lineas + 1
            
            v_socios = v_socios + 1
            v_socios1 = v_socios1 + 1
            
            Cad = "E6"
            Cad = Cad & RellenaABlancos(vsocio.NIF, True, 10)
            Cad = Cad & "T "
            Cad = Cad & Format(ActBanco, "0000")
            Cad = Cad & Format(v_Sucursal, "0000")
            Cad = Cad & RellenaABlancos(vsocio.CuentaBan, True, 10)
            Cad = Cad & Format(v_Provincia, "00")

            '[Monica]07/07/2014: si la factura es negativa ponemos un digito menos por el signo
            If Rs.Fields(3).Value < 0 Then
                Cad = Cad & Format(Round2(Rs.Fields(3).Value * 100, 0), "00000000000")
            Else
                Cad = Cad & Format(Round2(Rs.Fields(3).Value * 100, 0), "000000000000")
            End If

            Cad = Cad & RellenaABlancos(vsocio.Nombre, True, 36)
            Cad = Cad & RellenaABlancos(v_Domicilio, True, 36)
            Cad = Cad & RellenaABlancos(Trim(v_CPostal) & " " & v_Poblacion, True, 36)
            Cad = Cad & RellenaABlancos(vsocio.Digcontrol, True, 4)
            Cad = Cad & Format(v_lineas, "0000000")

            Print #NFich, Cad
            
            v_total = v_total + Rs.Fields(3).Value
            v_total1 = v_total1 + Rs.Fields(3).Value
            
            Set vsocio = Nothing
        End If
        
        Rs.MoveNext
    Wend
       
    v_lineas = v_lineas + 1
    
    Cad = RellenaABlancos("E8", True, 14)
    Cad = Cad & RellenaABlancos(Format(ActBanco, "0000"), True, 20)
    Cad = Cad & Format(Round2(v_total1 * 100, 0), "000000000000")
    Cad = Cad & Format(v_socios1, "0000000000")
    Cad = Cad & RellenaABlancos(Format(v_socios1 + 2, "0000000000"), True, 102)
    Cad = Cad & Format(v_lineas, "0000000")
    
    Print #NFich, Cad
    
    sql2 = "insert into tmpinformes (codusu, codigo1, importe1, importe2) values "
    sql2 = sql2 & "(" & vSesion.Codigo & "," & DBSet(ActBanco, "N") & "," & DBSet(v_socios1, "N")
    sql2 = sql2 & "," & DBSet(v_total1, "N") & ")"
    
    Conn.Execute sql2
    
    v_Entidades = v_Entidades + 1
    
    v_lineas = v_lineas + 1
    
    Cad = RellenaABlancos("E9", True, 34)
    Cad = Cad & Format(Round2(v_total * 100, 0), "000000000000")
    Cad = Cad & Format(v_socios, "0000000000")
    Cad = Cad & Format(v_lineas, "0000000000")
    
    Cad = Cad & RellenaABlancos(Format(v_Entidades, "000"), True, 92)
    Cad = Cad & Format(v_lineas, "0000000")
    
    Print #NFich, Cad
       
    Rs.Close
    Set Rs = Nothing
    
    Close (NFich)
    If v_socios > 0 Then GeneraFicheroNew = True
    Exit Function

EGen:
    Set Rs = Nothing
    Close (NFich)
    MuestraError Err.Number, Err.Description
End Function


Private Function GeneraFicheroNewAlz(vCadselect As String) As Boolean
Dim NFich As Integer
Dim Rs As ADODB.Recordset
Dim RsBanco As ADODB.Recordset
Dim Cad As String
Dim SQL As String
Dim I As Integer
Dim vsocio As CSocio
Dim v_total As Currency
Dim v_total1 As Currency
Dim v_total2 As Currency
Dim v_lineas As Currency
Dim v_socios As Currency
Dim v_socios1 As Currency
Dim v_socios2 As Currency
Dim v_dombanco As String
Dim v_pobbanco As String
Dim AntBanco As Integer
Dim ActBanco As Integer
Dim Banco As Currency
Dim sql2 As String

Dim v_Entidades As Integer

Dim v_Sucursal As Integer
Dim v_Domicilio As String
Dim v_CPostal As String
Dim v_Poblacion As String
Dim v_Provincia As Integer

Dim SqlBic As String
Dim RsBic As ADODB.Recordset
Dim Bic As String

Dim IBAN As String



    On Error GoTo EGen
    
    GeneraFicheroNewAlz = False

    SQL = "delete from tmpinformes where codusu = " & vSesion.Codigo
    Conn.Execute SQL

    NFich = FreeFile
    Open App.path & "\gasoleob.txt" For Output As #NFich

    Set Rs = New ADODB.Recordset
    
    SQL = vCadselect
    
    Rs.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    
    v_lineas = 0
    
    '***************REGISTRO R1
    v_lineas = v_lineas + 1
    
    Cad = "E1"
    Cad = Cad & RellenaABlancos(Format(txtCodigo(1).Text, "yymmdd"), True, 12)
    Cad = Cad & RellenaABlancos(Format(txtCodigo(2).Text, "0000"), True, 144)
    Cad = Cad & Format(v_lineas, "0000000")
    
    Print #NFich, Cad
    
    AntBanco = Rs.Fields(0).Value
    ActBanco = Rs.Fields(0).Value
    
    v_Entidades = 0
    
    v_total = 0  'importe total
    v_total1 = 0 'importe por banco
    v_socios = 0 ' socios total
    v_socios1 = 0 'socios por banco
    
    While Not Rs.EOF
        ActBanco = DBLet(Rs.Fields(0).Value, "N")
    
        If ActBanco <> AntBanco Or v_socios1 = 0 Then
            If v_socios1 <> 0 Then
                '***************REGISTRO R8
                v_lineas = v_lineas + 1
                
                Cad = RellenaABlancos("E8", True, 14)
                Cad = Cad & RellenaABlancos(Format(AntBanco, "0000"), True, 20)
                Cad = Cad & Format(Round2(v_total1 * 100, 0), "000000000000")
                Cad = Cad & Format(v_socios1, "0000000000")
                Cad = Cad & RellenaABlancos(Format(v_socios1 + 2, "0000000000"), True, 102)
                Cad = Cad & Format(v_lineas, "0000000")
                
                Print #NFich, Cad
                
                sql2 = "insert into tmpinformes (codusu, codigo1, importe1, importe2) values "
                sql2 = sql2 & "(" & vSesion.Codigo & "," & DBSet(AntBanco, "N") & "," & DBSet(v_socios1, "N")
                sql2 = sql2 & "," & DBSet(v_total1, "N") & ")"
                
                Conn.Execute sql2
                
                v_total1 = 0
                v_socios1 = 0
                
                v_Entidades = v_Entidades + 1
            End If
        
            '***************REGISTRO R2
            v_lineas = v_lineas + 1
            
            'Nro total de socios del banco
            v_total1 = 0
            
            Cad = RellenaABlancos("E2", True, 14)
            Cad = Cad & Format(txtCodigo(2).Text, "0000")
            
            Cad = Cad & RellenaABlancos(Format(ActBanco, "0000"), True, 140)
            Cad = Cad & Format(v_lineas, "0000000")
            Print #NFich, Cad
            
            AntBanco = ActBanco
        End If
    
        Set vsocio = New CSocio
        If vsocio.LeerDatos(DBLet(Rs.Fields(5).Value, "N")) Then
            ' datos del banco
            v_Sucursal = 0
            v_Domicilio = ""
            v_CPostal = ""
            v_Poblacion = ""
            v_Provincia = 0
            
            v_Sucursal = Rs.Fields(1).Value ' vsocio.Sucursal
            v_Domicilio = vsocio.Domicilio
            v_CPostal = vsocio.CPostal
            v_Poblacion = vsocio.POBLACION
            v_Provincia = Mid(vsocio.CPostal, 1, 2)
            
            
            Pb1.Value = Pb1.Value + 1
            DoEvents
            v_lineas = v_lineas + 1
            
            v_socios = v_socios + 1
            v_socios1 = v_socios1 + 1
            
'[Monica]19/01/2015: ahora pasan a ser con el Bic registros E4
'            Cad = "E6"
'            Cad = Cad & RellenaABlancos(vsocio.NIF, True, 10)
'            Cad = Cad & "T "
'            Cad = Cad & Format(Rs.Fields(0).Value, "0000")
'            Cad = Cad & Format(v_Sucursal, "0000")
'            Cad = Cad & RellenaABlancos(Rs.Fields(3).Value, True, 10)
'            Cad = Cad & Format(v_Provincia, "00")
'            Cad = Cad & Format(Round2(Rs.Fields(5).Value * 100, 0), "000000000000")
'
'            Cad = Cad & RellenaABlancos(vsocio.Nombre, True, 36)
'            Cad = Cad & RellenaABlancos(v_Domicilio, True, 36)
'            Cad = Cad & RellenaABlancos(Trim(v_CPostal) & " " & v_Poblacion, True, 36)
'            Cad = Cad & RellenaABlancos(Rs.Fields(2).Value, True, 4)
'            Cad = Cad & Format(v_lineas, "0000000")

            Cad = "E4"
            Cad = Cad & RellenaABlancos(vsocio.NIF, True, 9)
            Cad = Cad & "T"
            Cad = Cad & RellenaABlancos(Format(Rs.Fields(0).Value, "0000"), True, 11)
            Cad = Cad & RellenaABlancos(ReemplazaCharNoAdmitidos(vsocio.Nombre), True, 30)
            Cad = Cad & RellenaABlancos(ReemplazaCharNoAdmitidos(v_Domicilio), True, 30)
            Cad = Cad & RellenaABlancos(Trim(v_CPostal) & " " & ReemplazaCharNoAdmitidos(v_Poblacion), True, 31)
            
            '[Monica]07/07/2014: si la factura es negativa ponemos un digito menos por el signo
            If Rs.Fields(6).Value < 0 Then
                Cad = Cad & Format(Round2(Rs.Fields(6).Value * 100, 0), "000000000")
            Else
                Cad = Cad & Format(Round2(Rs.Fields(6).Value * 100, 0), "0000000000")
            End If
            
            
            IBAN = Rs.Fields(4) & Format(Rs.Fields(0), "0000") & Format(Rs.Fields(1).Value, "0000") & Format(Rs.Fields(2), "00") & Format(Rs.Fields(3), "0000000000")
            
            Cad = Cad & RellenaABlancos(IBAN, True, 34)
            Cad = Cad & Format(v_lineas, "0000000")

            Print #NFich, Cad
            
            v_total = v_total + Rs.Fields(6).Value
            v_total1 = v_total1 + Rs.Fields(6).Value
            
            Set vsocio = Nothing
        End If
        
        Rs.MoveNext
    Wend
       
    v_lineas = v_lineas + 1
    Cad = RellenaABlancos("E8", True, 14)
    Cad = Cad & RellenaABlancos(Format(ActBanco, "0000"), True, 20)
    Cad = Cad & Format(Round2(v_total1 * 100, 0), "000000000000")
    Cad = Cad & Format(v_socios1, "0000000000")
    Cad = Cad & RellenaABlancos(Format(v_socios1 + 2, "0000000000"), True, 102)
    Cad = Cad & Format(v_lineas, "0000000")
    
    Print #NFich, Cad
    
    sql2 = "insert into tmpinformes (codusu, codigo1, importe1, importe2) values "
    sql2 = sql2 & "(" & vSesion.Codigo & "," & DBSet(ActBanco, "N") & "," & DBSet(v_socios1, "N")
    sql2 = sql2 & "," & DBSet(v_total1, "N") & ")"
    
    Conn.Execute sql2
    
    v_Entidades = v_Entidades + 1
    
    v_lineas = v_lineas + 1
    
    Cad = RellenaABlancos("E9", True, 34)
    Cad = Cad & Format(Round2(v_total * 100, 0), "000000000000")
    Cad = Cad & Format(v_socios, "0000000000")
    Cad = Cad & Format(v_lineas, "0000000000")
    
    Cad = Cad & RellenaABlancos(Format(v_Entidades, "000"), True, 92)
    Cad = Cad & Format(v_lineas, "0000000")
    
    Print #NFich, Cad
       
    Rs.Close
    Set Rs = Nothing
    
    Close (NFich)
    If v_socios > 0 Then GeneraFicheroNewAlz = True
    Exit Function

EGen:
    Set Rs = Nothing
    Close (NFich)
    MuestraError Err.Number, Err.Description
End Function





'Funcion que devuelve que bancos no estan creados en entidaddom
Private Function ComprobarBancos(vSelect As String) As Boolean
Dim SQL As String

Dim Rs As ADODB.Recordset


    On Error GoTo eComprobarBancos
    
    ComprobarBancos = False
    
    SQL = "select distinct codbanco, codsucur from (" & vSelect & ") tabla "
    SQL = SQL & " where not (codbanco, codsucur)  in (select codentidad, codsucur from entidaddom) or codbanco is null or codsucur is null "
    Cadena = ""
    
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    While Not Rs.EOF
        Cadena = Cadena & "(" & DBLet(Rs!codbanco, "N") & "," & DBLet(Rs!codsucur, "N") & "), "
        
        Rs.MoveNext
    Wend
    Set Rs = Nothing
    
    If Cadena <> "" Then
        'quitamos la ultima coma
        Cadena = Mid(Cadena, 1, Len(Cadena) - 2)
    
        Set frmMens = New frmMensajes

        frmMens.OpcionMensaje = 23
        frmMens.cadWhere = SQL 'CADENA
        frmMens.Show vbModal

        Set frmMens = Nothing
        
        Exit Function
    End If
    
    ComprobarBancos = True
    Exit Function
    
eComprobarBancos:
    MuestraError Err.Number, "Comprobar Bancos", Err.Description
End Function




'Funcion que devuelve que bancos no estan creados en entidaddom
Private Function ComprobarBancosNew(vSelect As String) As Boolean
Dim SQL As String

Dim Rs As ADODB.Recordset


    On Error GoTo eComprobarBancos
    
    ComprobarBancosNew = False
    
    SQL = "select distinct codsocio, codbanco, codsucur from (" & vSelect & ") tabla "
    SQL = SQL & " where codbanco is null or codsucur is null or codbanco = 0 or codsucur = 0 "
    Cadena = ""
    
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    While Not Rs.EOF
        Cadena = Cadena & "(" & DBLet(Rs!codbanco, "N") & "," & DBLet(Rs!codsucur, "N") & "), "
        
        Rs.MoveNext
    Wend
    Set Rs = Nothing
    
    If Cadena <> "" Then
        'quitamos la ultima coma
        Cadena = Mid(Cadena, 1, Len(Cadena) - 2)
    
        Set frmMens = New frmMensajes

        frmMens.OpcionMensaje = 23
        frmMens.cadWhere = SQL 'CADENA
        frmMens.Show vbModal

        Set frmMens = Nothing
        
        Exit Function
    End If
    
    ComprobarBancosNew = True
    Exit Function
    
eComprobarBancos:
    MuestraError Err.Number, "Comprobar Bancos", Err.Description
End Function




Private Function ComprobarIBAN(vSelect As String) As Boolean
Dim SQL As String

Dim Rs As ADODB.Recordset
Dim cta As String
Dim CC As String
Dim DCcorrecto As String

    On Error GoTo eComprobarBancos
    
    ComprobarIBAN = False
    
    SQL = "select distinct codsocio, codbanco, codsucur, digcontr, cuentaba, iban from (" & vSelect & ") tabla "
    
    Cadena = ""
    
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    While Not Rs.EOF And Cadena = ""
        cta = Format(DBLet(Rs!codbanco), "0000") & Format(DBLet(Rs!codsucur), "0000") & Format(DBLet(Rs!digcontr), "00") & Format(DBLet(Rs!cuentaba), "0000000000")
        If Len(cta) = 20 Then
            
            DCcorrecto = DigitoControlCorrecto(cta)
            If DCcorrecto <> Format(DBLet(Rs!digcontr), "00") Then
                Cadena = "El socio " & DBLet(Rs!codsocio) & " tiene DC incorrecto, deberia ser " & DCcorrecto
            Else
                If DBLet(Rs!IBAN) = "" Then
                    'NO ha puesto IBAN
                    If DevuelveIBAN2("ES", cta, cta) Then Cadena = "El socio " & DBLet(Rs!codsocio) & " debería tener el iban ES" & cta
                Else
                    CC = CStr(Mid(DBLet(Rs!IBAN), 1, 2))
                    If DevuelveIBAN2(CStr(CC), cta, cta) Then
                        If Mid(DBLet(Rs!IBAN), 3) <> cta Then
                            Cadena = "El socio " & DBLet(Rs!codsocio) & " tiene IBAN distinto del calculado [" & CC & cta & "]"
                        End If
                    End If
                End If
            End If
        End If
    
        Rs.MoveNext
    Wend
    Set Rs = Nothing
    
    If Cadena <> "" Then
        ComprobarIBAN = False
        MsgBox Cadena, vbExclamation
        Exit Function
    End If
    
    ComprobarIBAN = True
    Exit Function
    
eComprobarBancos:
    MuestraError Err.Number, "Comprobar IBAN", Err.Description
End Function



Private Function ReemplazaCharNoAdmitidos(vCadena As String) As String
' ñ,Ñ,ç,Ç,º,ª,acentos
Dim vAux As String

    If vCadena = "" Then Exit Function
    
    vAux = vCadena
    
    vAux = Replace(vAux, "ñ", "n")
    vAux = Replace(vAux, "Ñ", "N")
    vAux = Replace(vAux, "ç", "c")
    vAux = Replace(vAux, "Ç", "C")
    vAux = Replace(vAux, "º", " ")
    vAux = Replace(vAux, "ª", " ")
    vAux = Replace(vAux, "á", "a")
    vAux = Replace(vAux, "é", "e")
    vAux = Replace(vAux, "í", "i")
    vAux = Replace(vAux, "ó", "o")
    vAux = Replace(vAux, "ú", "u")
    vAux = Replace(vAux, "Á", "A")
    vAux = Replace(vAux, "É", "E")
    vAux = Replace(vAux, "Í", "I")
    vAux = Replace(vAux, "Ó", "O")
    vAux = Replace(vAux, "Ú", "U")
    vAux = Replace(vAux, "'", " ")
    '[Monica]05/04/2017: faltaba la barra /
    vAux = Replace(vAux, "/", " ")
    
    
    ReemplazaCharNoAdmitidos = vAux
    
End Function
    
    
