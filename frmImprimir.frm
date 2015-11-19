VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmImprimir 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Impresión listados"
   ClientHeight    =   2910
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6780
   Icon            =   "frmImprimir.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2910
   ScaleWidth      =   6780
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ProgressBar pg1 
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   2340
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   6120
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdConfigImpre 
      Caption         =   "Sel. &impresora"
      Height          =   375
      Left            =   2400
      TabIndex        =   2
      Top             =   2340
      Width           =   1275
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   375
      Left            =   5340
      TabIndex        =   1
      Top             =   2340
      Width           =   1275
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "&Imprimir"
      Default         =   -1  'True
      Height          =   375
      Left            =   3900
      TabIndex        =   0
      Top             =   2340
      Width           =   1275
   End
   Begin VB.Frame Frame2 
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Width           =   6435
      Begin VB.CheckBox chkEMAIL 
         Caption         =   "Enviar e-mail"
         Height          =   195
         Left            =   4920
         TabIndex        =   8
         Top             =   180
         Width           =   1335
      End
      Begin VB.CheckBox chkSoloImprimir 
         Caption         =   "Previsualizar"
         Height          =   255
         Left            =   420
         TabIndex        =   5
         Top             =   180
         Width           =   1275
      End
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   180
      Locked          =   -1  'True
      TabIndex        =   3
      Text            =   "Sin definir"
      Top             =   180
      Width           =   6315
   End
   Begin VB.Image Image1 
      Height          =   495
      Left            =   240
      Top             =   1200
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   960
      TabIndex        =   7
      Top             =   1320
      Width           =   5535
   End
End
Attribute VB_Name = "frmImprimir"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' +-+-+-+-+-+-+-+-+-+-+-
' +-+- Autor: DAVID +-+-
' +-+-+-+-+-+-+-+-+-+-+-

Option Explicit

Public Opcion As Integer
'Equivale a OpcionListado en frmListado
    
'[Monica]01/08/2011: añadido el arigesmail
Public outCodigoCliProv As Long
Public outTipoDocumento As Byte
        '0 UNDEFINNED. Si es cero NO va por este trozo de programa
        '1.- Oferta cliente
        '2.- Pedido cliente
        '
        '
        'a partir del 50 van proveedores
        
        'a partir del 100 van socios

Public outClaveNombreArchiv As String  'Llevara el codigo oferta, pedido alb.....  SIN el .pdf, solo el nombre

    
    
    
Public FormulaSeleccion As String 'Formula de Seleccion para Crystal Report
Public SoloImprimir As Boolean
Public OtrosParametros As String   ' El grupo acaba en |
                                   ' param1=valor1|param2=valor2|
Public NumeroParametros As Integer
'Cuantos parametros hay.  EMPRESA(EMP) no es parametro. Es fijo en todos los informes

Public Titulo As String 'Titulo de descripción del informe
Public NombreRPT As String 'Nombre del Rpt


'Public NombreSubRptConta As String 'Nombre del subreport si va conectado a la BDatos Contabilidad
Public EnvioEMail As Boolean

'Public NumCod As String  'Codigo Traspaso, Movimiento Almacen
'Public Desde As Long
'Public Hasta As Long

Private MostrarTree As Boolean

Private MIPATH As String
Private Lanzado As Boolean
Private PrimeraVez As Boolean

Public ConSubInforme As Boolean 'Para saber si hay subinformes y hay que enlazar las
                                 'tablas a la BD correspondiente
Private InfConta As Boolean 'Para saber si el Informe se tiene que redireccionar
                            'a la BD de la Contabilidad de la Empresa
Private SubInformeConta As String






'Private ReestableceSoloImprimir As Boolean
Private Sub chkEMAIL_Click()
    If chkEMAIL.Value = 1 Then Me.chkSoloImprimir.Value = 0
End Sub

Private Sub chkSoloImprimir_Click()
    If Me.chkSoloImprimir.Value = 1 Then Me.chkEMAIL.Value = 0
End Sub


Private Sub cmdConfigImpre_Click()
    Screen.MousePointer = vbHourglass
    'Me.CommonDialog1.Flags = cdlPDPageNums
    CommonDialog1.ShowPrinter
    PonerNombreImpresora
    Screen.MousePointer = vbDefault
End Sub


Private Sub cmdImprimir_Click()
'Dim SQL As String
'Dim i As Long

    If Me.chkSoloImprimir.Value = 1 And Me.chkEMAIL.Value = 1 Then
        MsgBox "Si desea enviar por mail no debe marcar vista preliminar", vbExclamation
        Exit Sub
    End If
    
    Imprime
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub


Private Sub Form_Activate()
    If PrimeraVez Then
        espera 0.1
        '####Descomentar
'        CommitConexion
        If SoloImprimir Then
            Imprime
            Unload Me
        ElseIf Me.EnvioEMail Then
            Me.Hide
            DoEvents
            chkEMAIL.Value = 1
            Imprime
            Unload Me
        End If
    End If
    Screen.MousePointer = vbDefault
End Sub


Private Sub Form_Load()
Dim cad As String
Dim Nombre As String

    PrimeraVez = True
    Lanzado = False
    CargaICO
    cad = Dir(App.path & "\impre.dat", vbArchive)
    
    'ReestableceSoloImprimir = False
    If cad = "" Then
        chkSoloImprimir.Value = 0
    Else
        chkSoloImprimir.Value = 1
        'ReestableceSoloImprimir = True
    End If
    cmdImprimir.Enabled = True
    If SoloImprimir Then
        chkSoloImprimir.Value = 0
        Me.Frame2.Enabled = False
        chkSoloImprimir.visible = False
    Else
        Frame2.Enabled = True
        chkSoloImprimir.visible = True
    End If
    PonerNombreImpresora
    MostrarTree = False
    
    'A partir del infome 26, se trabajaba sobre la b de datos de informes(USUARIOS)

    MIPATH = App.path & "\Informes\"
    ConSubInforme = False
    InfConta = False
    
    Select Case Opcion

        Case 1
            ConSubInforme = True
            Text1.Text = Titulo
            
        Case 11
            SubInformeConta = "porciva.rpt"
            Text1.Text = Titulo
            Nombre = NombreRPT
            If Nombre = "rManArticResum.rpt" Then
                SubInformeConta = ""
            End If
        Case Else
            If Titulo <> "" Then
                Text1.Text = Titulo
                Nombre = NombreRPT
            Else
                Text1.Text = "Opcion incorrecta"
                Me.cmdImprimir.Enabled = False
            End If
    End Select
    If NombreRPT = "" Then NombreRPT = Nombre

    Screen.MousePointer = vbDefault
End Sub


Private Function Imprime() As Boolean
Dim Seguir As Boolean
Dim LanzaAbrirOutlook As Boolean


    Screen.MousePointer = vbHourglass
    CadenaDesdeOtroForm = ""
    With frmVisReport
        .FormulaSeleccion = Me.FormulaSeleccion
        .SoloImprimir = (Me.chkSoloImprimir.Value = 0)
        .OtrosParametros = OtrosParametros
        .NumeroParametros = NumeroParametros
        .MostrarTree = MostrarTree
        .Informe = MIPATH & NombreRPT
        .InfConta = InfConta
        
        
'        If NombreSubRptConta <> "" Then
'            .SubInformeConta = NombreSubRptConta
'        Else
'            .SubInformeConta = ""
'        End If
        .ConSubInforme = ConSubInforme
        .SubInformeConta = SubInformeConta
        .Opcion = Opcion
        .ExportarPDF = (chkEMAIL.Value = 1)
        .Show vbModal
    End With
    
    If Me.chkEMAIL.Value = 1 Then
    '####Descomentar
        If CadenaDesdeOtroForm <> "" Then
            
            If Me.EnvioEMail Then  'se llamo desde envio masivo
'                frmEMail.Show vbModal
                
            Else 'informe normal, pero que se selecciono enviar e-mail
                'Febrero 2010
                ' Nuevo
                LanzaAbrirOutlook = False
                If vParamAplic.ExeEnvioMail <> "" Then
                    If Me.outTipoDocumento = 0 Then
                        'MsgBox "Tipo de documento sin definir en el envio.", vbExclamation
                    Else
                        LanzaAbrirOutlook = True
                    End If
                End If
            
                If LanzaAbrirOutlook Then
                    '
                    LanzaProgramaAbrirOutlook
                Else
                    'El que habia
                    frmEMail.Opcion = 0
                    frmEMail.Show vbModal
                End If
            End If
            CadenaDesdeOtroForm = ""
        
        'frmEMail.Show vbModal
        End If
    End If
    Unload Me
  
End Function


Private Sub Form_Unload(Cancel As Integer)
    If Me.chkEMAIL.Value = 1 Then Me.chkSoloImprimir.Value = 1
    'If ReestableceSoloImprimir Then SoloImprimir = False
    OperacionesArchivoDefecto
'    NombreSubRptConta = ""
    outTipoDocumento = 0 'Para restear esta variable

End Sub

Private Sub OperacionesArchivoDefecto()
Dim crear  As Boolean
On Error GoTo ErrOperacionesArchivoDefecto

crear = (Me.chkSoloImprimir.Value = 1)
'crear = crear And ReestableceSoloImprimir
If Not crear Then
    Kill App.path & "\impre.dat"
    Else
        FileCopy App.path & "\Vacio.dat", App.path & "\impre.dat"
End If
ErrOperacionesArchivoDefecto:
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub Text1_DblClick()
Frame2.Tag = Val(Frame2.Tag) + 1
If Val(Frame2.Tag) > 2 Then
    Frame2.Enabled = True
    chkSoloImprimir.visible = True
End If
End Sub

Private Sub PonerNombreImpresora()
On Error Resume Next
    Label1.Caption = Printer.DeviceName
    If Err.Number <> 0 Then
        Label1.Caption = "No hay impresora instalada"
        Err.Clear
    End If
End Sub

Private Sub CargaICO()
    On Error Resume Next
'    Image1.Picture = LoadPicture(App.Path & "\iconos\printer.ico")
    Image1.Picture = frmPpal.imgListComun32.ListImages.Item(10).Picture
    If Err.Number <> 0 Then Err.Clear
End Sub


'[Monica]01/08/2011: añadido ARIMAILGES.EXE
Private Sub LanzaProgramaAbrirOutlook()
Dim NombrePDF As String
Dim Aux As String
Dim Lanza As String

    On Error GoTo ELanzaProgramaAbrirOutlook

    If Not PrepararCarpetasEnvioMail(True) Then Exit Sub

    'Primer tema. Copiar el docum.pdf con otro nombre mas significatiov
    Select Case outTipoDocumento
    Case 1
        'Oferta
        Aux = "OFE" & Me.outClaveNombreArchiv & ".pdf"
    Case 2
        'Fra
         Aux = Me.outClaveNombreArchiv & ".pdf"
    Case 3
         Aux = "PED" & Me.outClaveNombreArchiv & ".pdf"
    Case 4
         Aux = Me.outClaveNombreArchiv & ".pdf"
    Case 5
        Aux = "FPROF" & Me.outClaveNombreArchiv & ".pdf"
    
    Case 51
        Aux = "PEDP" & Me.outClaveNombreArchiv & ".pdf"
        
    Case 100
         Aux = Me.outClaveNombreArchiv & ".pdf"
        
    End Select
    NombrePDF = App.path & "\temp\" & Aux
    If Dir(NombrePDF, vbArchive) <> "" Then Kill NombrePDF
    FileCopy App.path & "\docum.pdf", NombrePDF
    
    Aux = FijaDireccionEmail
    Lanza = Aux & "|"
    Aux = ""
    Select Case outTipoDocumento
    Case 1
        Aux = "Oferta nº" & outClaveNombreArchiv
    Case 2
        Aux = "Factura nº" & outClaveNombreArchiv
    Case 3
        Aux = "Pedido cliente nº" & outClaveNombreArchiv
    Case 4
        Aux = "Albarán nº" & outClaveNombreArchiv
    Case 5
        Aux = "Factura proforma desde Oferta: " & outClaveNombreArchiv
        
        
    '--------------------------------------------------
    Case 51
        Aux = "Pedido proveedor nº: " & outClaveNombreArchiv
        
    Case 100
        Aux = "Factura nº" & outClaveNombreArchiv
        
    End Select
    
    Lanza = Lanza & Aux & "|"
    
    'Aqui pondremos lo del texto del BODY
    Aux = ""
    Lanza = Lanza & Aux & "|"
    
    
    'Envio o mostrar
    Lanza = Lanza & "0"   '0. Display   1.  send
    
    'Campos reservados para el futuro
    Lanza = Lanza & "||||"
    
    'El/los adjuntos
    Lanza = Lanza & NombrePDF & "|"
    
    Aux = App.path & "\" & vParamAplic.ExeEnvioMail & " " & Lanza
    Shell Aux, vbNormalFocus
    
    Exit Sub
ELanzaProgramaAbrirOutlook:
    MuestraError Err.Number, Err.Description
End Sub


Private Function FijaDireccionEmail() As String
Dim campoemail As String
Dim otromail As String


    FijaDireccionEmail = ""
    
    
    If outTipoDocumento < 50 Then
        'Clientes
        If outTipoDocumento = 1 Or outTipoDocumento = 2 Or outTipoDocumento = 3 Then
            campoemail = "maiclie1"
            otromail = "maiclie2"
        Else
            campoemail = "maiclie2"
            otromail = "maiclie1"
        End If
        campoemail = DevuelveDesdeBD(cPTours, campoemail, "scliente", "codclien", Me.outCodigoCliProv, "N") ' , otromail)
        If campoemail = "" Then campoemail = otromail
    Else
        If outTipoDocumento < 100 Then
            'Para provedores
            If outTipoDocumento = 51 Or outTipoDocumento = 52 Or outTipoDocumento = 53 Then
                campoemail = "maiprov1"
                otromail = "maiprov2"
            Else
                campoemail = "maiprov2"
                otromail = "maiprov1"
            End If
            campoemail = DevuelveDesdeBDNew(cPTours, campoemail, "sprove", "codprove", Me.outCodigoCliProv, "N", otromail)
            If campoemail = "" Then campoemail = otromail
        Else
            'Para Socios
            If outTipoDocumento >= 100 Then
                campoemail = "maisocio"
                otromail = "maisocio"
            Else
                campoemail = "maisocio"
                otromail = "maisocio"
            End If
            campoemail = DevuelveDesdeBDNew(cPTours, "ssocio", campoemail, "codsocio", Me.outCodigoCliProv, "N") ' , otromail)
            If campoemail = "" Then campoemail = otromail
        End If
    End If
    FijaDireccionEmail = campoemail
End Function


