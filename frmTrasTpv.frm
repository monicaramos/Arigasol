VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmTrasTpv 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Traspaso facturas TPV"
   ClientHeight    =   3585
   ClientLeft      =   45
   ClientTop       =   2430
   ClientWidth     =   4965
   Icon            =   "frmTrasTpv.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3585
   ScaleWidth      =   4965
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
      Height          =   3375
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   4755
      Begin MSComctlLib.ProgressBar Pb1 
         Height          =   255
         Left            =   600
         TabIndex        =   8
         Top             =   1920
         Width           =   3480
         _ExtentX        =   6138
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   3
         Left            =   1845
         MaxLength       =   10
         TabIndex        =   1
         Tag             =   "Código Postal|T|S|||clientes|codposta|||"
         Top             =   1200
         Width           =   1050
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   2
         Left            =   1845
         MaxLength       =   10
         TabIndex        =   0
         Tag             =   "Código Postal|T|S|||clientes|codposta|||"
         Top             =   840
         Width           =   1050
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   2985
         TabIndex        =   3
         Top             =   2520
         Width           =   975
      End
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   1800
         TabIndex        =   2
         Top             =   2520
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "Fecha Factura"
         ForeColor       =   &H00972E0B&
         Height          =   255
         Index           =   16
         Left            =   600
         TabIndex        =   7
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label Label4 
         Caption         =   "Desde"
         Height          =   195
         Index           =   15
         Left            =   960
         TabIndex        =   6
         Top             =   840
         Width           =   465
      End
      Begin VB.Label Label4 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   14
         Left            =   960
         TabIndex        =   5
         Top             =   1200
         Width           =   420
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   2
         Left            =   1530
         Picture         =   "frmTrasTpv.frx":000C
         ToolTipText     =   "Buscar fecha"
         Top             =   840
         Width           =   240
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   3
         Left            =   1530
         Picture         =   "frmTrasTpv.frx":0097
         ToolTipText     =   "Buscar fecha"
         Top             =   1200
         Width           =   240
      End
   End
End
Attribute VB_Name = "frmTrasTpv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' +-+-+-+-+-+-+-+-+-+-+-
' +-+- Autor MANOLO +-+-
' +-+-+-+-+-+-+-+-+-+-+-

Option Explicit

Private numser As String 'LETRA DE SERIE
Private TotalImp As Currency
    
Private Conexion As Byte
'1.- Conexión a BD Ariges  2.- Conexión a BD Conta

Private HaDevueltoDatos As Boolean

Private WithEvents frmC As frmCal 'calendario fechas
Attribute frmC.VB_VarHelpID = -1


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
Dim cDesde As String, cHasta As String 'cadena codigo Desde/Hasta
Dim nDesde As String, nHasta As String 'cadena Descripcion Desde/Hasta
Dim cadTABLA As String, cOrden As String
Dim i As Byte
    
    'D/H Fecha factura
    cDesde = Trim(txtCodigo(2).Text)
    cHasta = Trim(txtCodigo(3).Text)
    
    If DatosOk Then
        If TraspasoTPV(cDesde, cHasta) Then
            MsgBox "Traspaso a TPV realizado correctamente.", vbExclamation
            Pb1.visible = False
            cmdCancel_Click
        End If
    End If
    
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    If PrimeraVez Then
        PrimeraVez = False
        Pb1.visible = False
        PonerFoco txtCodigo(2)
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
    indFrame = 5
    tabla = "schfac"
            
    'Esto se consigue poneinedo el cancel en el opcion k corresponda
    Me.cmdCancel.Cancel = True
'    Me.Width = w + 70
'    Me.Height = h + 350
End Sub

Private Sub Form_Unload(Cancel As Integer)
    DesBloqueoManual ("TRASTPV")
End Sub

Private Sub frmC_Selec(vFecha As Date)
 'Fecha
    txtCodigo(CByte(imgFec(2).Tag)).Text = Format(vFecha, "dd/MM/yyyy")
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
'ahora
    If KeyAscii = teclaBuscar Then
        Select Case Index
            Case 2: KEYFecha KeyAscii, 2 'fecha desde
            Case 3: KEYFecha KeyAscii, 3 'fecha hasta
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
Dim cad As String, cadTipo As String 'tipo cliente

    'Quitar espacios en blanco por los lados
    txtCodigo(Index).Text = Trim(txtCodigo(Index).Text)
    
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub

    
    Select Case Index
            
        Case 2, 3 'FECHAS
            If txtCodigo(Index).Text <> "" Then PonerFormatoFecha txtCodigo(Index)
            
    End Select
End Sub

Private Sub FrameCobrosVisible(visible As Boolean, ByRef h As Integer, ByRef w As Integer)
    Me.FrameCobros.visible = visible
End Sub

Private Function TraspasoTPV(desde As String, hasta As String) As Boolean
    Dim sql As String
    Dim SqlWhere As String
    Dim RS As ADODB.Recordset
    Dim importel As Currency
    Dim impbase As Currency
    
    Dim actFactura As Integer
    Dim antfactura As Integer
    
    Dim TotalReg As Currency
    Dim db As BaseDatos
    
    TraspasoTPV = False
    MensError = ""
    sql = "SELECT count(*) from scaalb " & _
            " where fecalbar >= " & DBSet(desde, "F") & _
            " and fecalbar <= " & DBSet(hasta, "F") & _
            " and numfactu <> 0 "
    
    TotalReg = TotalRegistros(sql)
    
    If TotalReg = 0 Then
        MsgBox "No existen valores entre estos límites. Reintroduzca.", vbExclamation
    Else
        Pb1.visible = True
        Pb1.Max = TotalReg
        sql = "SELECT scaalb.*, sartic.codigiva, sartic.impuesto from scaalb, sartic " & _
                " where fecalbar >= " & DBSet(txtCodigo(2), "F") & _
                " and fecalbar <= " & DBSet(txtCodigo(3), "F") & _
                " and numfactu <> 0 and scaalb.codartic=sartic.codartic" & _
                " order by scaalb.numfactu, scaalb.numlinea"
                
        Set db = New BaseDatos
        db.abrir vSesion.CadenaConexion, "root", "aritel"
        db.Tipo = "MYSQL"
                
        If TraspasoHistoricoFacturas(db, sql, txtCodigo(2).Text, txtCodigo(3).Text, Pb1) Then TraspasoTPV = True
        
        Set db = Nothing
    End If
    
End Function

Private Function DatosOk() As Boolean
Dim cDesde As String
Dim cHasta  As String
Dim ctipo As String

    DatosOk = True
    cDesde = txtCodigo(2).Text
    cHasta = txtCodigo(3).Text
    
    If cDesde = "" Or cHasta = "" Then
        MsgBox "Los campos de Fecha han de tener un valor.", vbExclamation
        DatosOk = False
        Exit Function
    End If
    
    If Not EsFechaOK(cDesde) Or Not EsFechaOK(cHasta) Then
        DatosOk = False
        Exit Function
    End If
    
    If Not EsFechaIgualPosterior(cDesde, cHasta, True, "La Fecha Desde ha de ser inferior a la Fecha Hasta. Reintroduzca.") Then
        DatosOk = False
        Exit Function
    End If
    
    ' comprobamos que no exista ninguna linea de albaran ya introducida en el historico de lineas de factura
    ctipo = ""
    ctipo = DevuelveDesdeBD("letraser", "stipom", "codtipom", "FAT", "T")
    If ctipo = "" Then
        MsgBox "No existe el tipo de movimiento de Traspaso a TPV", vbExclamation
        DatosOk = False
        Exit Function
    End If
    
    If ExisteEnHistorico(cDesde, cHasta, ctipo) Then
        MsgBox "Hay líneas de albaranes que ya están en el histórico de facturas. Revise.", vbExclamation
        DatosOk = False
        Exit Function
    End If
    
End Function

