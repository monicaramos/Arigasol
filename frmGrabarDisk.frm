VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmGrabarDisk 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Grabación Fichero de Tarjetas Gasolinera"
   ClientHeight    =   4470
   ClientLeft      =   45
   ClientTop       =   2430
   ClientWidth     =   7185
   Icon            =   "frmGrabarDisk.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4470
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
      Height          =   4215
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   6915
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Index           =   3
         Left            =   1770
         MaxLength       =   20
         TabIndex        =   4
         Top             =   2910
         Width           =   4065
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   450
         Top             =   3480
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   6
         Left            =   2685
         Locked          =   -1  'True
         TabIndex        =   14
         Text            =   "Text5"
         Top             =   2520
         Width           =   3135
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   6
         Left            =   1770
         MaxLength       =   2
         TabIndex        =   3
         Top             =   2520
         Width           =   830
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   2
         Left            =   1785
         MaxLength       =   10
         TabIndex        =   2
         Tag             =   "Código Postal|T|S|||clientes|codposta|||"
         Top             =   1920
         Width           =   1050
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   4905
         TabIndex        =   6
         Top             =   3510
         Width           =   975
      End
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   3720
         TabIndex        =   5
         Top             =   3510
         Width           =   975
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   0
         Left            =   1785
         MaxLength       =   6
         TabIndex        =   0
         Top             =   840
         Width           =   855
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   1
         Left            =   1800
         MaxLength       =   6
         TabIndex        =   1
         Top             =   1170
         Width           =   830
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   2685
         Locked          =   -1  'True
         TabIndex        =   9
         Text            =   "Text5"
         Top             =   840
         Width           =   3135
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   2685
         Locked          =   -1  'True
         TabIndex        =   8
         Text            =   "Text5"
         Top             =   1170
         Width           =   3135
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Texto "
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
         Left            =   540
         TabIndex        =   16
         Top             =   2940
         Width           =   465
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   6
         Left            =   1485
         MouseIcon       =   "frmGrabarDisk.frx":000C
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar colectivo"
         Top             =   2520
         Width           =   240
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
         Left            =   540
         TabIndex        =   15
         Top             =   2520
         Width           =   660
      End
      Begin VB.Label Label4 
         Caption         =   "Fecha "
         ForeColor       =   &H00972E0B&
         Height          =   255
         Index           =   16
         Left            =   540
         TabIndex        =   13
         Top             =   1920
         Width           =   765
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   2
         Left            =   1485
         Picture         =   "frmGrabarDisk.frx":015E
         ToolTipText     =   "Buscar fecha"
         Top             =   1920
         Width           =   240
      End
      Begin VB.Label Label4 
         Caption         =   "Desde"
         Height          =   195
         Index           =   13
         Left            =   960
         TabIndex        =   12
         Top             =   840
         Width           =   465
      End
      Begin VB.Label Label4 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   12
         Left            =   960
         TabIndex        =   11
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
         Left            =   540
         TabIndex        =   10
         Top             =   600
         Width           =   495
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   0
         Left            =   1485
         MouseIcon       =   "frmGrabarDisk.frx":01E9
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar cliente"
         Top             =   840
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   1
         Left            =   1485
         MouseIcon       =   "frmGrabarDisk.frx":033B
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar cliente"
         Top             =   1215
         Width           =   240
      End
   End
End
Attribute VB_Name = "frmGrabarDisk"
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
Dim i As Byte
Dim sql As String

    If DatosOk Then
        sql = "select count(*) from soltarje where 1 = 1 "
        If txtCodigo(0).Text <> "" Then sql = sql & " and codsocio >= " & DBSet(txtCodigo(0).Text, "N")
        If txtCodigo(1).Text <> "" Then sql = sql & " and codsocio <= " & DBSet(txtCodigo(1).Text, "N")
        sql = sql & " and codsocio in (select codsocio from ssocio where codcoope = " & DBSet(txtCodigo(6).Text, "N") & ")"
    
        If RegistrosAListar(sql) <> 0 Then
            If GeneraFichero Then
                If CopiarFichero Then
                    If MsgBox("¿ Desea eliminar registros ?", vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
                        
                        sql = "delete from " & tabla & " where 1 = 1 "
                        If txtCodigo(0).Text <> "" Then sql = sql & " and codsocio >= " & DBSet(txtCodigo(0).Text, "N")
                        If txtCodigo(1).Text <> "" Then sql = sql & " and codsocio <= " & DBSet(txtCodigo(1).Text, "N")
                        sql = sql & " and codsocio in (select codsocio from ssocio where codcoope = " & DBSet(txtCodigo(6).Text, "N") & ")"
                        
                        Conn.Execute sql
                    End If
                    Unload Me
                End If
            End If
        Else
            MsgBox "No hay registros para generar el fichero", vbExclamation
        End If
    End If
    
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
    Limpiar Me

    'IMAGES para busqueda
     Me.imgBuscar(0).Picture = frmPpal.imgListImages16.ListImages(1).Picture
     Me.imgBuscar(1).Picture = frmPpal.imgListImages16.ListImages(1).Picture
     Me.imgBuscar(6).Picture = frmPpal.imgListImages16.ListImages(1).Picture

    '###Descomentar
'    CommitConexion

    FrameCobrosVisible True, h, w
    indFrame = 5
    tabla = "soltarje"
    txtCodigo(3).Text = DevuelveDesdeBDNew(cPTours, "sempre", "nomempre", "codempre", 1, "N")

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
        Case 6 ' Colectivo
            AbrirFrmColectivo (Index)
    End Select
    PonerFoco txtCodigo(indCodigo)
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
            Case 0: KEYBusqueda KeyAscii, 0 'cliente desde
            Case 1: KEYBusqueda KeyAscii, 1 'cliente hasta
            Case 6: KEYBusqueda KeyAscii, 6 'colectivo
            Case 2: KEYFecha KeyAscii, 2 'fecha desde
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

        Case 2 'FECHA
            If txtCodigo(Index).Text <> "" Then PonerFormatoFecha txtCodigo(Index)

        Case 6 'COLECTIVO
            txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), "scoope", "nomcoope", "codcoope", "N")
            If txtCodigo(Index).Text <> "" Then txtCodigo(Index).Text = Format(txtCodigo(Index).Text, "00")
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
    indCodigo = indice
    Set frmCol = New frmManCoope
    frmCol.DatosADevolverBusqueda = "0|1|"
    frmCol.DeConsulta = True
    frmCol.CodigoActual = txtCodigo(indCodigo)
    frmCol.Show vbModal
    Set frmCol = Nothing
End Sub

Private Function GeneraFichero() As Boolean
Dim NFich As Integer
Dim Regs As Integer
Dim CodigoOrdenante As String
Dim Importe As Currency
Dim RS As ADODB.Recordset
Dim RT As ADODB.Recordset
Dim Aux As String
Dim cad As String
Dim sql As String
Dim SQL1 As String
Dim vsufijo As String
Dim CuentaPropia As String
Dim i As Integer
Dim Tarjeta As String

    On Error GoTo EGen
    GeneraFichero = False

    NFich = FreeFile
    Open App.path & "\temp.txt" For Output As #NFich

    Set RS = New ADODB.Recordset
    
    'partimos de la tabla soltarje
    sql = "SELECT soltarje.*, ssocio.nomsocio from soltarje, ssocio where 1 = 1 "
    If txtCodigo(0).Text <> "" Then sql = sql & " and soltarje.codsocio >= " & DBSet(txtCodigo(0).Text, "N")
    If txtCodigo(1).Text <> "" Then sql = sql & " and soltarje.codsocio <= " & DBSet(txtCodigo(1).Text, "N")

    sql = sql & " and soltarje.codsocio = ssocio.codsocio "
    sql = sql & " and ssocio.codcoope = " & DBSet(txtCodigo(6).Text, "N")

    RS.Open sql, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    
    Regs = 0
    
    While Not RS.EOF

        Regs = Regs + 1
        For i = 1 To RS!Numtarje
            cad = "97240000"
            
            Tarjeta = ""
            Tarjeta = DevuelveDesdeBDNew(cPTours, "starje", "numtarje", "codsocio", RS!codsocio, "N", , "tiptarje", RS!tiptarje, "N")
            cad = cad & Mid(Tarjeta, 1, 8)
'            cad = cad & Format(RS!Numtarje, "00000000")

            Select Case RS!tiptarje
                Case 0  'Normal
                    cad = cad & RellenaABlancos(txtCodigo(3).Text, True, 20)
                    cad = cad & Space(27)
                    cad = cad & RellenaABlancos(RS!NomSocio, True, 27)
                    cad = cad & Space(27)
                    cad = cad & Space(27)
                    Print #NFich, cad
                Case 1  'Bonificada
                    cad = cad & "GASOLEO BONIFICADO  "
                    cad = cad & Space(27)
                    cad = cad & RellenaABlancos(RS!NomSocio, True, 27)
                    cad = cad & Space(27)
                    cad = cad & RellenaABlancos(txtCodigo(3).Text, True, 27)
                    Print #NFich, cad
            End Select
        Next i
        
        RS.MoveNext
    Wend
       
    RS.Close
    Set RS = Nothing
    
    Close (NFich)
    If Regs > 0 Then GeneraFichero = True
    Exit Function
EGen:
    Set RS = Nothing
    Close (NFich)
    MuestraError Err.Number, Err.Description

End Function


Public Function CopiarFichero() As Boolean
Dim nomfich As String
Dim cadena As String
On Error GoTo ecopiarfichero

    CopiarFichero = False
    ' abrimos el commondialog para indicar donde guardarlo
    Me.CommonDialog1.InitDir = App.path
    Me.CommonDialog1.DefaultExt = "lum"
    cadena = Format(CDate(txtCodigo(2).Text), FormatoFecha)
    CommonDialog1.Filter = "Archivos lum|lum|"
    CommonDialog1.FilterIndex = 1
    
    CommonDialog1.FileName = "1" & Format(txtCodigo(6).Text, "0") & Mid(cadena, 3, 2) & Mid(cadena, 6, 2) & Mid(cadena, 9, 2)
    
    Me.CommonDialog1.ShowSave
    If CommonDialog1.FileName <> "" Then
        FileCopy App.path & "\temp.txt", CommonDialog1.FileName
        CopiarFichero = True
    End If

ecopiarfichero:
    If Err.Number <> 0 Then
        MuestraError Err.Number, Err.Description
    End If
    Err.Clear

End Function

Private Function RellenaABlancos(cadena As String, PorLaDerecha As Boolean, longitud As Integer) As String
Dim cad As String
    
    cad = Space(longitud)
    If PorLaDerecha Then
        cad = cadena & cad
        RellenaABlancos = Left(cad, longitud)
    Else
        cad = cad & cadena
        RellenaABlancos = Right(cad, longitud)
    End If
    
End Function

Private Function DatosOk() As Boolean

    DatosOk = True
    
    If txtCodigo(2).Text = "" Then
        MsgBox "Debe introducir un valor en el campo fecha.", vbExclamation
        PonerFoco txtCodigo(2)
        DatosOk = False
        Exit Function
    End If
    
    If txtCodigo(6).Text = "" Then
        MsgBox "Debe introducir una cooperativa.", vbExclamation
        PonerFoco txtCodigo(6)
        DatosOk = False
        Exit Function
    End If
    
    
End Function
