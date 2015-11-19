VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmCambioCliente 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cambio de Cliente"
   ClientHeight    =   5640
   ClientLeft      =   45
   ClientTop       =   2430
   ClientWidth     =   8160
   Icon            =   "frmCambioCliente.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5640
   ScaleWidth      =   8160
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameSeleccion 
      Height          =   5505
      Left            =   30
      TabIndex        =   16
      Top             =   30
      Width           =   8055
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFC0C0&
         Height          =   315
         Index           =   1
         Left            =   6270
         TabIndex        =   21
         Text            =   "Text2"
         Top             =   4980
         Width           =   1305
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFC0C0&
         Height          =   315
         Index           =   0
         Left            =   4710
         TabIndex        =   20
         Text            =   "Text1"
         Top             =   4980
         Width           =   1485
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   1
         Left            =   1500
         TabIndex        =   19
         Top             =   4980
         Width           =   1095
      End
      Begin VB.CommandButton CmdAceptar 
         Caption         =   "Aceptar"
         Height          =   375
         Index           =   1
         Left            =   240
         TabIndex        =   18
         Top             =   4980
         Width           =   1185
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   4425
         Left            =   150
         TabIndex        =   17
         Top             =   420
         Width           =   7725
         _ExtentX        =   13626
         _ExtentY        =   7805
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         Checkboxes      =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   16777215
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.Image Image1 
         Height          =   240
         Index           =   1
         Left            =   390
         Picture         =   "frmCambioCliente.frx":000C
         ToolTipText     =   "Desmarcar todos"
         Top             =   120
         Width           =   240
      End
      Begin VB.Image Image1 
         Height          =   240
         Index           =   0
         Left            =   150
         Picture         =   "frmCambioCliente.frx":0A0E
         ToolTipText     =   "Marcar todos"
         Top             =   120
         Width           =   240
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00C00000&
         X1              =   4650
         X2              =   7560
         Y1              =   4890
         Y2              =   4890
      End
      Begin VB.Label Label1 
         Caption         =   "Totales"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   1
         Left            =   3630
         TabIndex        =   22
         Top             =   5010
         Width           =   795
      End
   End
   Begin VB.Frame FrameCobros 
      Height          =   5415
      Left            =   60
      TabIndex        =   8
      Top             =   60
      Width           =   6375
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   2
         Left            =   1680
         MaxLength       =   2
         TabIndex        =   1
         Top             =   1170
         Width           =   1065
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   2
         Left            =   2820
         Locked          =   -1  'True
         TabIndex        =   32
         Text            =   "Text5"
         Top             =   1170
         Width           =   3135
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   0
         Left            =   1680
         MaxLength       =   10
         TabIndex        =   0
         Tag             =   "Código Postal|T|S|||clientes|codposta|||"
         Top             =   360
         Width           =   1050
      End
      Begin VB.Frame Frame2 
         Caption         =   "Datos para la contabilización"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   1455
         Left            =   300
         TabIndex        =   25
         Top             =   2790
         Width           =   5625
         Begin VB.TextBox txtNombre 
            BackColor       =   &H80000018&
            Enabled         =   0   'False
            Height          =   285
            Index           =   1
            Left            =   2550
            TabIndex        =   28
            Top             =   300
            Width           =   2985
         End
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   1
            Left            =   1860
            MaxLength       =   10
            TabIndex        =   4
            Tag             =   "Cta.Contable|T|S|||sparam|ctaconta|||"
            Top             =   300
            Width           =   585
         End
         Begin VB.TextBox txtNombre 
            BackColor       =   &H80000018&
            Enabled         =   0   'False
            Height          =   285
            Index           =   4
            Left            =   2550
            TabIndex        =   27
            Top             =   660
            Width           =   2985
         End
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   4
            Left            =   1860
            MaxLength       =   10
            TabIndex        =   5
            Tag             =   "Cta.Contable|T|S|||sparam|ctaconta|||"
            Top             =   660
            Width           =   585
         End
         Begin VB.TextBox txtNombre 
            BackColor       =   &H80000018&
            Enabled         =   0   'False
            Height          =   285
            Index           =   5
            Left            =   2550
            TabIndex        =   26
            Top             =   1050
            Width           =   2985
         End
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   5
            Left            =   1860
            MaxLength       =   10
            TabIndex        =   6
            Tag             =   "Cta.Dif.negativas|T|S|||sparam|ctanegtat|||"
            Top             =   1050
            Width           =   585
         End
         Begin VB.Label Label1 
            Caption         =   "Número Diario "
            ForeColor       =   &H00972E0B&
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   31
            Top             =   360
            Width           =   1350
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   1
            Left            =   1530
            ToolTipText     =   "Buscar Diario"
            Top             =   330
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Concepto al Haber"
            ForeColor       =   &H00972E0B&
            Height          =   195
            Index           =   25
            Left            =   120
            TabIndex        =   30
            Top             =   1050
            Width           =   1410
         End
         Begin VB.Label Label1 
            Caption         =   "Concepto al Debe"
            ForeColor       =   &H00972E0B&
            Height          =   195
            Index           =   24
            Left            =   120
            TabIndex        =   29
            Top             =   690
            Width           =   1350
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   4
            Left            =   1530
            ToolTipText     =   "Buscar Concepto"
            Top             =   690
            Width           =   240
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   5
            Left            =   1530
            ToolTipText     =   "Buscar Concepto"
            Top             =   1050
            Width           =   240
         End
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   1680
         TabIndex        =   24
         Top             =   750
         Width           =   1095
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   3
         Left            =   1695
         MaxLength       =   10
         TabIndex        =   3
         Tag             =   "Código Postal|T|S|||clientes|codposta|||"
         Top             =   2310
         Width           =   1050
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   6
         Left            =   1695
         MaxLength       =   10
         TabIndex        =   2
         Tag             =   "Código Postal|T|S|||clientes|codposta|||"
         Top             =   1950
         Width           =   1050
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   0
         Left            =   4890
         TabIndex        =   9
         Top             =   4800
         Width           =   975
      End
      Begin VB.CommandButton CmdAceptar 
         Caption         =   "&Aceptar"
         Height          =   375
         Index           =   0
         Left            =   3720
         TabIndex        =   7
         Top             =   4800
         Width           =   975
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   0
         Left            =   2820
         Locked          =   -1  'True
         TabIndex        =   10
         Text            =   "Text5"
         Top             =   360
         Width           =   3075
      End
      Begin MSComctlLib.ProgressBar Pb1 
         Height          =   255
         Left            =   450
         TabIndex        =   15
         Top             =   4440
         Width           =   5430
         _ExtentX        =   9578
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Nueva F.Pago"
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
         Left            =   360
         TabIndex        =   33
         Top             =   1170
         Width           =   1020
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   2
         Left            =   1410
         MouseIcon       =   "frmCambioCliente.frx":7260
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar F.Pago"
         Top             =   1170
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Tarjeta"
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
         Index           =   0
         Left            =   360
         TabIndex        =   23
         Top             =   750
         Width           =   525
      End
      Begin VB.Label Label4 
         Caption         =   "Fecha Ticket"
         ForeColor       =   &H00972E0B&
         Height          =   255
         Index           =   16
         Left            =   360
         TabIndex        =   14
         Top             =   1620
         Width           =   1815
      End
      Begin VB.Label Label4 
         Caption         =   "Desde"
         Height          =   195
         Index           =   15
         Left            =   810
         TabIndex        =   13
         Top             =   1950
         Width           =   465
      End
      Begin VB.Label Label4 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   14
         Left            =   810
         TabIndex        =   12
         Top             =   2310
         Width           =   420
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   6
         Left            =   1410
         Picture         =   "frmCambioCliente.frx":73B2
         ToolTipText     =   "Buscar fecha"
         Top             =   1950
         Width           =   240
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   3
         Left            =   1410
         Picture         =   "frmCambioCliente.frx":743D
         ToolTipText     =   "Buscar fecha"
         Top             =   2310
         Width           =   240
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
         Left            =   360
         TabIndex        =   11
         Top             =   360
         Width           =   495
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   0
         Left            =   1380
         MouseIcon       =   "frmCambioCliente.frx":74C8
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar cliente"
         Top             =   360
         Width           =   240
      End
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   7680
      Top             =   5160
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmCambioCliente"
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

Private WithEvents frmCol As frmManCoope 'Colectivo
Attribute frmCol.VB_VarHelpID = -1
Private WithEvents frmcli As frmManClien 'Clientes
Attribute frmcli.VB_VarHelpID = -1
Private WithEvents frmC As frmCal 'calendario fechas
Attribute frmC.VB_VarHelpID = -1
Private WithEvents frmFpa As frmManFpago 'mantenimiento de formas de pago
Attribute frmFpa.VB_VarHelpID = -1
Private WithEvents frmTDia As frmDiaConta
Attribute frmTDia.VB_VarHelpID = -1
Private WithEvents frmConce As frmConceConta 'conceptos de contabilidad
Attribute frmConce.VB_VarHelpID = -1

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
Dim NRegSelec As Integer
Dim CtaNuevoCliente As String



Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, 0, cerrar
    If cerrar Then Unload Me
End Sub

Private Sub cmdAceptar_Click(Index As Integer)
Dim cDesde As String, cHasta As String 'cadena codigo Desde/Hasta
Dim nDesde As String, nHasta As String 'cadena Descripcion Desde/Hasta
Dim cadTABLA As String, cOrden As String
Dim cadMen As String
Dim i As Byte
Dim sql As String
Dim tipo As Byte
Dim nRegs As Integer
Dim NumError As Long
Dim db As BaseDatos


    Select Case Index
        Case 0
            If Not DatosOk Then Exit Sub
            
            If CargarListView = 0 Then
                VisualizarListview False
                MsgBox "No existen datos entre esos límites.", vbExclamation
            Else
                VisualizarListview True
            End If
        Case 1
            If NRegSelec = 0 Then
                MsgBox "No ha seleccionado ningún ticket para realizar cambio de cliente.", vbExclamation
                Exit Sub
            Else
                If DatosContablesRequeridos Then
                    MsgBox "Ha de introducir los datos contables", vbExclamation
                    VisualizarListview False
                    PonerFoco txtCodigo(1)
                    Exit Sub
                Else
                    If MsgBox("Desea continuar con el proceso de cambio de cliente.", vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
                        If ProcesarCambios Then
                           MsgBox "Proceso realizado correctamente", vbExclamation
                           cmdCancel_Click (0)
                        End If
                    End If
                End If
            End If
    End Select

End Sub

Private Sub cmdCancel_Click(Index As Integer)
    Unload Me
End Sub

Private Sub Form_Activate()
    If PrimeraVez Then
        PrimeraVez = False
        ValoresPorDefecto
        PonerFoco txtCodigo(0)
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
Dim h As Integer, w As Integer
Dim List As Collection
Dim i As Integer

    PrimeraVez = True
    limpiar Me

    VisualizarListview False
    
    NRegSelec = 0
    
    'IMAGES para busqueda
    Me.imgBuscar(0).Picture = frmPpal.imgListImages16.ListImages(1).Picture
    Me.imgBuscar(1).Picture = frmPpal.imgListImages16.ListImages(1).Picture
    Me.imgBuscar(2).Picture = frmPpal.imgListImages16.ListImages(1).Picture
    Me.imgBuscar(4).Picture = frmPpal.imgListImages16.ListImages(1).Picture
    Me.imgBuscar(5).Picture = frmPpal.imgListImages16.ListImages(1).Picture


    '###Descomentar
'    CommitConexion
         
    FrameCobrosVisible True, h, w
    indFrame = 5
    tabla = "scaalb"
    
    Pb1.visible = False
    
    'Esto se consigue poneinedo el cancel en el opcion k corresponda
    Me.cmdCancel(0).Cancel = True
    Me.Width = w + 70
    Me.Height = h + 350
End Sub

Private Sub Form_Unload(Cancel As Integer)
    DesBloqueoManual ("FACTURAC")
End Sub

Private Sub frmC_Selec(vFecha As Date)
 'Fecha
    txtCodigo(CByte(imgFec(6).Tag)).Text = Format(vFecha, "dd/MM/yyyy")
End Sub

Private Sub frmCli_DatoSeleccionado(CadenaSeleccion As String)
'Form de Consulta de Clientes
    txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000000")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmFpa_DatoSeleccionado(CadenaSeleccion As String)
'Form de Consulta de Formas de Pago
    txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "00")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmTDia_DatoSeleccionado(CadenaSeleccion As String)
'Form de Consulta de tipos de diarios
    txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "00")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmConce_DatoSeleccionado(CadenaSeleccion As String)
'Form de Consulta de Clientes
    txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub Image1_Click(Index As Integer)
Dim i As Integer
Dim TotalCant As Currency
Dim TotalImporte As Currency

    Screen.MousePointer = vbHourglass
    
    TotalCant = 0
    TotalImporte = 0
    NRegSelec = 0
    
    Select Case Index
        Case 0
            For i = 1 To ListView1.ListItems.Count
                ListView1.ListItems(i).Checked = True
                TotalCant = TotalCant + CCur(ListView1.ListItems(i).SubItems(3))
                TotalImporte = TotalImporte + CCur(ListView1.ListItems(i).SubItems(4))
                NRegSelec = NRegSelec + 1
            Next i
        Case 1
            For i = 1 To ListView1.ListItems.Count
                ListView1.ListItems(i).Checked = False
            Next i

    End Select
    Screen.MousePointer = vbDefault

    Text1(0).Text = TotalCant
    Text1(1).Text = TotalImporte

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
    imgFec(6).Tag = Index 'independentment de les dates que tinga, sempre pose l'index en la 27
    If txtCodigo(Index).Text <> "" Then frmC.NovaData = txtCodigo(Index).Text

    frmC.Show vbModal
    Set frmC = Nothing
    If imgFec(6).Tag = 6 Then
        PonerFoco txtCodigo(3)
    Else
        PonerFoco txtCodigo(1)
    End If
    ' ***************************
End Sub

Private Sub ListView1_ItemCheck(ByVal item As MSComctlLib.ListItem)
Dim i As Integer
Dim TotalCant As Currency
Dim TotalImporte As Currency
    
    Screen.MousePointer = vbHourglass
    
    TotalCant = 0
    TotalImporte = 0
    NRegSelec = 0
    
    ' vemos si lo podemos seleccionar
    For i = 1 To ListView1.ListItems.Count
        If ListView1.ListItems(i).Checked Then
            TotalCant = TotalCant + CCur(ListView1.ListItems(i).SubItems(3))
            TotalImporte = TotalImporte + CCur(ListView1.ListItems(i).SubItems(4))
            NRegSelec = NRegSelec + 1
        End If
    Next i
    
    Screen.MousePointer = vbDefault

    Text1(0).Text = TotalCant
    Text1(1).Text = TotalImporte
    
End Sub

Private Sub ListView1_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub imgBuscar_Click(Index As Integer)
   Select Case Index
        Case 0 'CLIENTE
            AbrirFrmClientes (Index)
        Case 1 'diario
            AbrirFrmDiario (Index)
        Case 2 ' forma de pago
            AbrirFrmFpagos (Index)
        Case 4, 5 'CONCEPTOS CONTABLES
            AbrirFrmConceptos (Index)
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
'14/02/2007
'    KEYpress KeyAscii
'ahora
    If KeyAscii = teclaBuscar Then
        Select Case Index
            Case 0: KEYBusqueda KeyAscii, 0 'cliente desde
            Case 1: KEYBusqueda KeyAscii, 1 'numero de diario
            Case 2: KEYBusqueda KeyAscii, 2 'forma de pago
            Case 4: KEYBusqueda KeyAscii, 4 'concepto al debe
            Case 5: KEYBusqueda KeyAscii, 5 'concepto al haber
            Case 6: KEYFecha KeyAscii, 6 'fecha desde
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
            
        Case 0 'CLIENTE
            txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), "ssocio", "nomsocio", "codsocio", "N")
            If txtCodigo(Index).Text <> "" Then txtCodigo(Index).Text = Format(txtCodigo(Index).Text, "000000")
        
            CargaCombo
        
        Case 1 'DIARIO DE CONTABILIDAD
            If txtCodigo(Index).Text <> "" Then
                txtNombre(Index).Text = ""
                txtNombre(Index).Text = DevuelveDesdeBDNew(cConta, "tiposdiario", "desdiari", "numdiari", txtCodigo(Index).Text, "N")
                If txtNombre(Index).Text = "" Then
                    MsgBox "Número de Diario no existe en la contabilidad. Reintroduzca.", vbExclamation
                    PonerFoco txtCodigo(Index)
                End If
            End If
            
        Case 2 ' FORMA DE PAGO
            txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), "sforpa", "nomforpa", "codforpa", "N") 'VRS:2.0.1(4)
            If txtCodigo(Index).Text <> "" Then txtCodigo(Index).Text = Format(txtCodigo(Index).Text, "00")
            
        Case 6, 3 'FECHAS
            If txtCodigo(Index).Text <> "" Then PonerFormatoFecha txtCodigo(Index)
        
        Case 4, 5 'CONCEPTOS
            If txtCodigo(Index).Text <> "" Then
                txtNombre(Index).Text = PonerNombreConcepto(txtCodigo(Index))
                If txtNombre(Index).Text = "" Then
                    MsgBox "Número de Concepto no existe en la contabilidad. Reintroduzca.", vbExclamation
                    PonerFoco txtCodigo(Index)
                End If
            End If
        
    End Select
End Sub

Private Sub FrameCobrosVisible(visible As Boolean, ByRef h As Integer, ByRef w As Integer)
    Me.FrameCobros.visible = visible
    If visible = True Then
        Me.FrameCobros.Top = -90
        Me.FrameCobros.Left = 0
        Me.FrameCobros.Height = 6015
        Me.FrameCobros.Width = 6555
        w = Me.FrameCobros.Width
        h = Me.FrameCobros.Height
    End If
End Sub

Private Sub ValoresPorDefecto()
   ' txtCodigo(6).Text = Format(Now, "dd/mm/yyyy")
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


Private Sub AbrirFrmDiario(indice As Integer)
    indCodigo = indice
    Set frmTDia = New frmDiaConta
    frmTDia.DatosADevolverBusqueda = "0|1|"
    frmTDia.CodigoActual = txtCodigo(indCodigo)
    frmTDia.Show vbModal
    Set frmTDia = Nothing
End Sub

Private Sub AbrirFrmConceptos(indice As Integer)
    indCodigo = indice
    Set frmConce = New frmConceConta
    frmConce.DatosADevolverBusqueda = "0|1|"
    frmConce.CodigoActual = txtCodigo(indCodigo)
    frmConce.Show vbModal
    Set frmConce = Nothing
End Sub

Private Sub AbrirFrmFpagos(indice As Integer)
    indCodigo = indice
    Set frmFpa = New frmManFpago
    frmFpa.DatosADevolverBusqueda = "0|1|"
    frmFpa.DeConsulta = True
    frmFpa.CodigoActual = txtCodigo(indCodigo)
    frmFpa.Show vbModal
    Set frmFpa = Nothing
End Sub

Private Function DatosOk() As Boolean
Dim b As Boolean
Dim sql As String

    b = True
    
    If vParamAplic.CtaContable <> "" Then
        sql = ""
        sql = DevuelveDesdeBD("codsocio", "ssocio", "codmacta", vParamAplic.CtaContable, "T")
        If sql = "" Then
            MsgBox "No existe un cliente de contado. Revise", vbExclamation
            b = False
            DatosOk = b
            Exit Function
        End If
    End If

    If txtCodigo(0).Text = "" Then
        MsgBox "Debe introducir obligatoriamente un cliente.", vbExclamation
        b = False
        DatosOk = b
        Exit Function
    Else
        CtaNuevoCliente = ""
        CtaNuevoCliente = DevuelveDesdeBD("codmacta", "ssocio", "codsocio", txtCodigo(0).Text, "N")
        If CtaNuevoCliente = "" Then
            MsgBox "El nuevo cliente no tiene una cuenta contable existente. Revise.", vbExclamation
            b = False
            DatosOk = b
            Exit Function
        Else
            ' vemos si la cuenta contable existe en contabilidad
            sql = ""
            sql = DevuelveDesdeBDNew(cConta, "cuentas", "codmacta", "codmacta", CtaNuevoCliente, "T")
            If sql = "" Then
                MsgBox "La Cuenta Contable del cliente no existe en la Contabilidad", vbExclamation
                b = False
                DatosOk = b
                Exit Function
            End If
        End If
    End If

    ' comprobamos que la forma de pago de no admite bonificaciones
    If txtCodigo(2).Text = "" Then
        MsgBox "Debe introducir un valor para la Forma de Pago", vbExclamation
        b = False
        DatosOk = b
        Exit Function
    Else
        sql = ""
        sql = DevuelveDesdeBD("tipforpa", "sforpa", "codforpa", txtCodigo(2).Text, "N")
        '[Monica]04/01/2013: efectivos
        If sql <> "0" And sql <> "6" Then
            MsgBox "Debe introducir una Forma de Pago de Efectivo. Reintroduzca.", vbExclamation
            b = False
            DatosOk = b
            Exit Function
        Else
            If AdmiteBonificacion(txtCodigo(2).Text) Then
                MsgBox "Debe introducir una Forma de Pago que no permita bonificación. Reintroduzca.", vbExclamation
                b = False
                DatosOk = b
                Exit Function
            End If
        End If
    End If
    
    DatosOk = b
    
End Function

Private Function CargarListView() As Integer
'Muestra la lista Detallada de Facturas que dieron error al contabilizar
'en un ListView
Dim RS As ADODB.Recordset
Dim ItmX As ListItem
Dim sql As String
Dim HayReg As Integer

    On Error GoTo ECargarList
    
    HayReg = 0
    CargarListView = 0
    If ListView1.ListItems.Count <> 0 Then Exit Function

    Screen.MousePointer = vbHourglass

    Me.FrameSeleccion.Height = 5415
    Me.FrameSeleccion.Width = 8055
    Me.Height = 6120
    Me.Width = 8370
    

    sql = " SELECT  scaalb.numalbar, scaalb.fecalbar, sartic.nomartic, scaalb.cantidad, scaalb.importel, scaalb.codclave, scaalb.codsocio, scaalb.codturno "
    sql = sql & " FROM scaalb, sartic, ssocio where ssocio.codmacta = " & DBSet(vParamAplic.CtaContable, "T")
'    SQL = SQL & " and scaalb.codforpa = " & DBSet(txtCodigo(2).Text, "N")
    If txtCodigo(6).Text <> "" Then sql = sql & " and fecalbar >= " & DBSet(txtCodigo(6).Text, "F")
    If txtCodigo(3).Text <> "" Then sql = sql & " and fecalbar <= " & DBSet(txtCodigo(3).Text, "F")
    sql = sql & " and scaalb.numfactu = 0 "
    sql = sql & " and scaalb.declaradogp = 0 " ' que no haya sido declarado como gp
    sql = sql & " and scaalb.codsocio = ssocio.codsocio  and scaalb.codartic = sartic.codartic"
    sql = sql & " order by scaalb.fecalbar, scaalb.numalbar "
    
    Set RS = New ADODB.Recordset
    RS.Open sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not RS.EOF Then
        VisualizarListview True
    
        'Los encabezados
        
        ListView1.ColumnHeaders.Clear

        ListView1.ColumnHeaders.Add , , "Ticket", 1100
        ListView1.ColumnHeaders.Add , , "Fecha", 1100, 1
        ListView1.ColumnHeaders.Add , , "Nombre Articulo", 2500
        ListView1.ColumnHeaders.Add , , "Cantidad", 1300, 1
        ListView1.ColumnHeaders.Add , , "Importe", 1400, 1
        ListView1.ColumnHeaders.Add , , "Nreg", 0
        ListView1.ColumnHeaders.Add , , "Socio", 0
        ListView1.ColumnHeaders.Add , , "Turno", 0
       
        ListView1.ListItems.Clear
        
        While Not RS.EOF
            Set ItmX = ListView1.ListItems.Add
            'El primer campo será codtipom si llamamos desde Ventas
            ' y será codprove si llamamos desde Compras
            ItmX.Text = Format(RS!numalbar, "00000000")
            ItmX.SubItems(1) = DBLet(RS!fecAlbar, "F")
            ItmX.SubItems(2) = DBLet(RS!NomArtic, "T")
            ItmX.SubItems(3) = DBLet(RS!cantidad, "N")
            ItmX.SubItems(4) = DBLet(RS!importel, "N")
            ItmX.SubItems(5) = DBLet(RS!Codclave, "N")
            ItmX.SubItems(6) = DBLet(RS!codsocio, "N")
            ItmX.SubItems(7) = DBLet(RS!codTurno, "N")
            HayReg = 1
            RS.MoveNext
        Wend
    End If
    RS.Close
    Set RS = Nothing
    CargarListView = HayReg
    
    Screen.MousePointer = vbDefault
    
ECargarList:
    If Err.Number <> 0 Then
        Err.Clear
    End If
End Function

Private Sub VisualizarListview(Modo As Boolean)
    If Modo = False Then
        Me.Width = 6570
    Else
        Me.Width = 8250
    End If
    FrameSeleccion.visible = Modo
    FrameCobros.visible = Not Modo
End Sub

Private Sub CargaCombo()
Dim sql As String
Dim RS As ADODB.Recordset

    Combo1.Clear
    
    sql = "select * from starje where codsocio = " & DBSet(txtCodigo(0).Text, "N")
    
    Set RS = New ADODB.Recordset
    RS.Open sql, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    
    'Combo1.AddItem "" 'pose uno en blanc sinse valor
    While Not RS.EOF
        Combo1.AddItem RS!Numtarje
        Combo1.ItemData(Combo1.NewIndex) = RS!NumLinea
        RS.MoveNext
    Wend
    
    Combo1.Text = Combo1.List(0)

    RS.Close
    Set RS = Nothing
    
ErrCarga:
    If Err.Number <> 0 Then MuestraError Err.Number, "Cargar datos combo.", Err.Description
End Sub
    
Private Function ProcesarCambios() As Boolean
Dim sql As String
Dim SQL1 As String
Dim i As Integer
Dim HayReg As Integer
Dim b As Boolean

On Error GoTo eProcesarCambios

    
    sql = "update scaalb set codsocio = " & DBSet(txtCodigo(0).Text, "N") & ", numtarje = " & DBSet(Combo1.Text, "N")
    sql = sql & ", codforpa = " & DBSet(txtCodigo(2).Text, "N")
    HayReg = 0
    
    VisualizarListview False
    b = True
    If CrearTMPAsiento Then
        
        Conn.BeginTrans
        
        For i = 1 To ListView1.ListItems.Count
            If ListView1.ListItems(i).Checked Then
               
                ' si esta la contabilizacion del cierre de turno tengo que deshacerla
                ' realizando un asiento que introduzco en el diario
                If EstaContabilizadoCierre(ListView1.ListItems(i).SubItems(5)) Then
                    HayReg = 1
'                    Debug.Print ListView1.ListItems(i).SubItems(1)
                    InsertaLineaEnTemporal ListView1.ListItems(i)
                End If
                
                SQL1 = sql & " where codclave = " & ListView1.ListItems(i).SubItems(5)
                Conn.Execute SQL1
            
            End If
        Next i
        If HayReg = 1 Then b = InsertarAsientoContabilidad
    End If
'    conn.RollbackTrans
    Conn.CommitTrans
    
    BorrarTMPAsiento
    
    ProcesarCambios = True

eProcesarCambios:
    If Err.Number <> 0 Or Not b Then
        Conn.RollbackTrans
        ProcesarCambios = False
    End If
End Function

Private Function EstaContabilizadoCierre(NumReg As Long) As Boolean
Dim RS As ADODB.Recordset
Dim Rs1 As ADODB.Recordset
Dim sql As String
Dim Sql2 As String

    EstaContabilizadoCierre = False

    sql = "select * from scaalb where codclave = " & DBSet(NumReg, "N")
    
    Set RS = New ADODB.Recordset
    RS.Open sql, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    
    If Not RS.EOF Then
        Sql2 = "SELECT count(*)" & _
              " FROM srecau " & _
              "WHERE srecau.fechatur = " & DBSet(RS!fecAlbar, "F") & " and " & _
                   " srecau.codturno = " & DBSet(RS!codTurno, "N") & " and " & _
                   " srecau.codforpa = " & DBSet(RS!Codforpa, "N") & " and " & _
                   " srecau.intconta = 1 "
                   
        Set Rs1 = New ADODB.Recordset
        Rs1.Open Sql2, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    
        If Not Rs1.EOF Then
            If Rs1.Fields(0).Value <> 0 Then EstaContabilizadoCierre = True
        End If
        Set Rs1 = Nothing
        
    End If
    Set RS = Nothing
End Function


Private Sub InsertaLineaEnTemporal(ByRef ItmX As ListItem)
Dim sql As String
Dim Codmacta As String
Dim RS As ADODB.Recordset
Dim SQL1 As String

        Codmacta = ""
        Codmacta = DevuelveDesdeBD("codmacta", "ssocio", "codsocio", ItmX.SubItems(6), "N")
        
        SQL1 = "insert into tmpasien(fecalbar, codturno, codmacta, importel) values ("
        SQL1 = SQL1 & DBSet(ItmX.SubItems(1), "F") & ","
        SQL1 = SQL1 & DBSet(ItmX.SubItems(7), "N") & ","
        SQL1 = SQL1 & DBSet(Codmacta, "T") & ","
        SQL1 = SQL1 & DBSet(ItmX.SubItems(4), "N") & ")"

        Conn.Execute SQL1
    
    Set RS = Nothing
    
End Sub


Private Function InsertarAsientoContabilidad() As Boolean
Dim RS As ADODB.Recordset
Dim sql As String
Dim Mc As CContadorContab
Dim AntFec As String
Dim AntTur As String
Dim ActFec As String
Dim ActTur As String
Dim NumLinea As Integer
Dim Obs As String
Dim cadMen As String
Dim b As Boolean
Dim i As Integer
Dim ImporteD As Currency
Dim ImporteH As Currency
Dim Diferencia As Currency
Dim cad As String
Dim numdocum As String
Dim ampliacion As String
Dim ampliaciond As String
Dim ampliacionh As String

    
    On Error GoTo eInsertarAsientoContabilidad
     
    ConnConta.BeginTrans
    
    sql = "select count(distinct fecalbar, codmacta) from tmpasien "
    
    NumLinea = TotalRegistros(sql)
    NumLinea = NumLinea + 1
    Me.Pb1.visible = True
    
    If NumLinea > 0 Then
    
        sql = "select fecalbar, codmacta, sum(importel) from tmpasien " & _
                  "group by fecalbar, codmacta " & _
                  "order by fecalbar, codmacta "
        
        Set RS = New ADODB.Recordset
        RS.Open sql, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
        
        CargarProgres Me.Pb1, NumLinea

        AntFec = RS!fecAlbar
'        AntTur = RS!codTurno
        Set Mc = New CContadorContab
        
        If Mc.ConseguirContador("0", (CDate(RS!fecAlbar) <= CDate(FFin)), True) = 0 Then
        
            Obs = "Regularización Cierre Turno de fecha " & Format(RS!fecAlbar, "dd/mm/yyyy") '& " y turno T-" & Format(RS!codTurno, "0")
            
            ActFec = RS!fecAlbar
            
            While Not RS.EOF
                ActFec = RS!fecAlbar
                If AntFec <> ActFec Then 'Or AntTur <> RS!codTurno Then
                    
                        Obs = "Cambio Cierre Turno de fecha " & Format(AntFec, "dd/mm/yyyy") '& " y turno T-" & Format(RS!codTurno, "0")
                    
                        'Insertar en la conta Cabecera Asiento
                        b = InsertarCabAsientoDia(txtCodigo(1), Mc.Contador, AntFec, Obs, cadMen)
                        cadMen = "Insertando Cab. Asiento: " & cadMen
                        
                        'Insertar la linea de diferencias en el haber del cliente nuevo
                        i = i + 1
                        
                        If ImporteD > ImporteH Then
                            Diferencia = ImporteD - ImporteH
                        Else
                            Diferencia = ImporteH - ImporteD
                        End If
                        
                        cad = DBSet(txtCodigo(1).Text, "N") & "," & DBSet(AntFec, "F") & "," & DBSet(Mc.Contador, "N") & ","
                        cad = cad & DBSet(i, "N") & "," & DBSet(CtaNuevoCliente, "T") & "," & DBSet(numdocum, "T") & ","
                        
                        If ImporteD < ImporteH Then
                            cad = cad & DBSet(txtCodigo(4).Text, "N") & "," & DBSet(ampliacion, "T") & ","
                            cad = cad & DBSet(Diferencia, "N") & "," & ValorNulo & ","
                        Else
                            cad = cad & DBSet(txtCodigo(5).Text, "N") & "," & DBSet(ampliacion, "T") & ","
                            cad = cad & ValorNulo & "," & DBSet(Diferencia, "N") & ","
                        End If
                        
                        cad = cad & ValorNulo & "," & ValorNulo & "," & ValorNulo & ",0"
                        
                        cad = "(" & cad & ")"
                    
                        b = InsertarLinAsientoDia(cad, cadMen)
                        cadMen = "Insertando Lin. Asiento: " & i
            
                        IncrementarProgres Me.Pb1, 1
                        Me.Refresh
                        
                        AntFec = RS!fecAlbar
'                        AntTur = RS!codTurno
                        
                        Mc.ConseguirContador "0", (CDate(RS!fecAlbar) <= CDate(FFin)), True
                End If
                
                ImporteD = 0
                ImporteH = 0
                
                numdocum = Format(CDate(RS!fecAlbar), "ddmmyy") '& "-T" & Format(RS!codTurno, "0")
                ampliacion = "C.Turno " & Format(RS!fecAlbar, "dd/mm/yy") '& " T-" & Format(RS!codTurno, "0")
                ampliaciond = Trim(DevuelveDesdeBDNew(cConta, "conceptos", "nomconce", "codconce", txtCodigo(4).Text, "N")) & " " & ampliacion
                ampliacionh = Trim(DevuelveDesdeBDNew(cConta, "conceptos", "nomconce", "codconce", txtCodigo(5).Text, "N")) & " " & ampliacion
            
                i = i + 1
                
                cad = DBSet(txtCodigo(1).Text, "N") & "," & DBSet(RS!fecAlbar, "F") & "," & DBSet(Mc.Contador, "N") & ","
                cad = cad & DBSet(i, "N") & "," & DBSet(RS!Codmacta, "T") & "," & DBSet(numdocum, "T") & ","
                
                ' COMPROBAMOS EL SIGNO DEL IMPORTE PQ NO PERMITIMOS INTRODUCIR APUNTES CON IMPORTES NEGATIVOS
                If RS.Fields(2).Value > 0 Then
                    ' importe al debe en positivo
                    cad = cad & DBSet(txtCodigo(4).Text, "N") & "," & DBSet(ampliaciond, "T") & "," & DBSet(RS.Fields(2).Value, "N") & ","
                    cad = cad & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & ",0"
                
                    ImporteD = ImporteD + CCur(RS.Fields(2).Value)
                Else
                    ' importe al haber en positivo, cambiamos el signo
                    cad = cad & DBSet(txtCodigo(5).Text, "N") & "," & DBSet(ampliacionh, "T") & "," & ValorNulo & "," & DBSet((RS.Fields(2).Value * -1), "N") & "," & ValorNulo & ","
                    cad = cad & ValorNulo & "," & ValorNulo & ",0"
                
                    ImporteH = ImporteH + (CCur(RS.Fields(2).Value) * (-1))
                End If
                
                cad = "(" & cad & ")"
                
                b = InsertarLinAsientoDia(cad, cadMen)
                cadMen = "Insertando Lin. Asiento: " & i
            
                IncrementarProgres Me.Pb1, 1
                Me.Refresh
                
                RS.MoveNext
                
               ' ActTur = RS!codTurno
            Wend

        
            Obs = "Cambio Cierre Turno de fecha " & Format(ActFec, "dd/mm/yyyy") '& " y turno T-" & Format(AntTur, "0")
        
            'Insertar en la conta Cabecera Asiento
            b = InsertarCabAsientoDia(txtCodigo(1).Text, Mc.Contador, CStr(ActFec), Obs, cadMen)
            cadMen = "Insertando Cab. Asiento: " & cadMen
            
            'Insertar la linea de diferencias en el haber del cliente nuevo
            i = i + 1
            
            If ImporteD > ImporteH Then
                Diferencia = ImporteD - ImporteH
            Else
                Diferencia = ImporteH - ImporteD
            End If
            
            cad = DBSet(txtCodigo(1).Text, "N") & "," & DBSet(ActFec, "F") & "," & DBSet(Mc.Contador, "N") & ","
            cad = cad & DBSet(i, "N") & "," & DBSet(CtaNuevoCliente, "T") & "," & DBSet(numdocum, "T") & ","
            
            If ImporteD < ImporteH Then
                cad = cad & DBSet(txtCodigo(4).Text, "N") & "," & DBSet(ampliaciond, "T") & ","
                cad = cad & DBSet(Diferencia, "N") & "," & ValorNulo & ","
            Else
                cad = cad & DBSet(txtCodigo(5).Text, "N") & "," & DBSet(ampliacionh, "T") & ","
                cad = cad & ValorNulo & "," & DBSet(Diferencia, "N") & ","
            End If
            
            cad = cad & ValorNulo & "," & ValorNulo & "," & ValorNulo & ",0"
            
            cad = "(" & cad & ")"
        
            b = InsertarLinAsientoDia(cad, cadMen)
            cadMen = "Insertando Lin. Asiento: " & i
    
            IncrementarProgres Me.Pb1, 1
    '        Me.lblProgres(1).Caption = "Insertando línea de Asiento en Contabilidad...   (" & i & " de " & NumLinea & ")"
            Me.Refresh
            

        End If

    Set RS = Nothing
    
    End If
eInsertarAsientoContabilidad:
    If Err.Number <> 0 Then
        MsgBox "Error en la contabilización de asientos.", vbExclamation
        InsertarAsientoContabilidad = False
        ConnConta.RollbackTrans
    Else
        InsertarAsientoContabilidad = True
        ConnConta.CommitTrans
    End If
End Function


Private Function DatosContablesRequeridos() As Boolean
Dim i As Integer
Dim RS As ADODB.Recordset

    DatosContablesRequeridos = False
    If txtCodigo(1).Text = "" Or txtCodigo(4).Text = "" Or txtCodigo(5).Text = "" Then
        For i = 1 To ListView1.ListItems.Count
            If ListView1.ListItems(i).Checked Then
                ' si esta la contabilizacion del cierre de turno tengo que deshacerla
                ' realizando un asiento que introduzco en el diario
                If EstaContabilizadoCierre(ListView1.ListItems(i).SubItems(5)) Then
                    DatosContablesRequeridos = True
                    Exit Function
                End If
            End If
        Next i
    End If
End Function
