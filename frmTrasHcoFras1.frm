VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmTrasHcoFras1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Traspaso Histórico Facturas 1"
   ClientHeight    =   3585
   ClientLeft      =   45
   ClientTop       =   2430
   ClientWidth     =   6735
   Icon            =   "frmTrasHcoFras1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3585
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
      Height          =   3375
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   6585
      Begin MSComctlLib.ProgressBar Pb1 
         Height          =   255
         Left            =   360
         TabIndex        =   5
         Top             =   2250
         Width           =   5850
         _ExtentX        =   10319
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   3
         Left            =   2775
         MaxLength       =   10
         TabIndex        =   0
         Tag             =   "Código Postal|T|S|||clientes|codposta|||"
         Top             =   1740
         Width           =   1050
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   5205
         TabIndex        =   2
         Top             =   2730
         Width           =   975
      End
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   4020
         TabIndex        =   1
         Top             =   2730
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "Inserción de Datos"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   3
         Left            =   360
         TabIndex        =   9
         Top             =   2520
         Width           =   1725
      End
      Begin VB.Label Label4 
         Caption         =   "No debe haber nadie trabajando en la aplicación"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   2
         Left            =   330
         TabIndex        =   8
         Top             =   1140
         Width           =   5145
      End
      Begin VB.Label Label4 
         Caption         =   "Se recomienda hacer una copia antes de realizar el traspaso"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   1
         Left            =   330
         TabIndex        =   7
         Top             =   720
         Width           =   6015
      End
      Begin VB.Label Label4 
         Caption         =   "Se va a proceder a traspasar las facturas al Histórico 1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   0
         Left            =   330
         TabIndex        =   6
         Top             =   300
         Width           =   5895
      End
      Begin VB.Label Label4 
         Caption         =   "Hasta Fecha Factura"
         ForeColor       =   &H00972E0B&
         Height          =   255
         Index           =   16
         Left            =   600
         TabIndex        =   4
         Top             =   1740
         Width           =   1725
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   3
         Left            =   2490
         Picture         =   "frmTrasHcoFras1.frx":000C
         ToolTipText     =   "Buscar fecha"
         Top             =   1740
         Width           =   240
      End
   End
End
Attribute VB_Name = "frmTrasHcoFras1"
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
    cHasta = Trim(txtcodigo(3).Text)
    
    If DatosOk Then
        If TraspasoHistorico1(cHasta) Then
            MsgBox "Traspaso a Histórico 1 realizado correctamente.", vbExclamation
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
        If Day(Now) = 31 And Month(Now) = 12 Then
            txtcodigo(3).Text = "31/12/" & Format(Year(Now) - 5, "0000")
        Else
            txtcodigo(3).Text = "31/12/" & Format(Year(Now) - 6, "0000")
        End If
        PonerFoco txtcodigo(3)
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
            
    Me.Label4(3).visible = False
            
            
    'Esto se consigue poneinedo el cancel en el opcion k corresponda
    Me.cmdCancel.Cancel = True
'    Me.Width = w + 70
'    Me.Height = h + 350
End Sub

Private Sub Form_Unload(Cancel As Integer)
    DesBloqueoManual ("TRASHCO")
End Sub

Private Sub frmC_Selec(vFecha As Date)
 'Fecha
    txtcodigo(CByte(imgFec(3).Tag)).Text = Format(vFecha, "dd/MM/yyyy")
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
    imgFec(3).Tag = Index 'independentment de les dates que tinga, sempre pose l'index en la 27
    If txtcodigo(Index).Text <> "" Then frmC.NovaData = txtcodigo(Index).Text

    frmC.Show vbModal
    Set frmC = Nothing
    PonerFoco txtcodigo(CByte(imgFec(3).Tag) + 2)
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
    ConseguirFoco txtcodigo(Index), 3
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
    txtcodigo(Index).Text = Trim(txtcodigo(Index).Text)
    
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub

    
    Select Case Index
            
        Case 3 'FECHAS
            If txtcodigo(Index).Text <> "" Then PonerFormatoFecha txtcodigo(Index)
            
    End Select
End Sub

Private Sub FrameCobrosVisible(visible As Boolean, ByRef h As Integer, ByRef w As Integer)
    Me.FrameCobros.visible = visible
End Sub

Private Function TraspasoHistorico1(hasta As String) As Boolean
    Dim Sql As String
    Dim Sql2 As String
    Dim SqlWhere As String
    Dim Rs As ADODB.Recordset
    Dim importel As Currency
    Dim impbase As Currency
    
    Dim actFactura As Integer
    Dim antfactura As Integer
    
    Dim TotalReg As Currency
    Dim db As BaseDatos
    
    On Error GoTo eTraspasoHistorico1
    
    
    TraspasoHistorico1 = False
    
    MensError = ""
    
    Sql = "SELECT count(*) from schfac " & _
            " where fecfactu <= " & DBSet(hasta, "F")
    
    TotalReg = TotalRegistros(Sql)
    If TotalReg = 0 Then
        MsgBox "No existen datos a traspasar. Reintroduzca.", vbExclamation
        Exit Function
    End If
    
    Screen.MousePointer = vbHourglass
    
    Conn.BeginTrans
    
    Label4(3).Caption = "Inserción de Datos"
    Label4(3).visible = True
    DoEvents
    
    Sql = "SELECT * from schfac " & _
            " where fecfactu <= " & DBSet(hasta, "F")
    
    ' cabecera
    Sql2 = "insert into schfac1 "
    
    Conn.Execute Sql2 & Sql
    
    ' lineas
    Sql = "SELECT * from slhfac " & _
            " where fecfactu <= " & DBSet(hasta, "F")
    
    Sql2 = "insert into slhfac1 "
    
    Conn.Execute Sql2 & Sql
    
    ' vencimientos
    Sql = "SELECT * from svenci " & _
            " where fecfactu <= " & DBSet(hasta, "F")
    
    Sql2 = "insert into svenci1 "
    
    Conn.Execute Sql2 & Sql
    
    
    Label4(3).Caption = "Borrado de datos"
    DoEvents
    
    Sql = "delete from schfac where fecfactu <= " & DBSet(hasta, "F")
    Conn.Execute Sql
    
    Sql = "delete from slhfac where fecfactu <= " & DBSet(hasta, "F")
    Conn.Execute Sql
    
    Sql = "delete from svenci where fecfactu <= " & DBSet(hasta, "F")
    Conn.Execute Sql
    
    
    TraspasoHistorico1 = True
    Conn.CommitTrans
    Label4(3).visible = False
    Screen.MousePointer = vbDefault
    Exit Function
    
    
eTraspasoHistorico1:
    Label4(3).visible = False
    Screen.MousePointer = vbDefault
    Conn.RollbackTrans
    MuestraError Err.Number, "Traspaso Histórico 1", Err.Description
End Function

Private Function DatosOk() As Boolean
Dim cDesde As String
Dim cHasta  As String
Dim ctipo As String

    DatosOk = True
    
    cHasta = txtcodigo(3).Text
    
    If cHasta = "" Then
        MsgBox "El campo de Fecha ha de tener un valor.", vbExclamation
        DatosOk = False
        Exit Function
    End If
    
    If Not EsFechaOK(cHasta) Then
        DatosOk = False
        Exit Function
    End If
    
    
End Function

