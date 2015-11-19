VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmTraspasoTPV 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Trapaso Histórico de Facturas TPV"
   ClientHeight    =   3645
   ClientLeft      =   45
   ClientTop       =   2430
   ClientWidth     =   7185
   Icon            =   "frmTraspasoTPV.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3645
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
      Height          =   3405
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   6915
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   3
         Left            =   1815
         MaxLength       =   10
         TabIndex        =   1
         Tag             =   "Código Postal|T|S|||clientes|codposta|||"
         Top             =   1290
         Width           =   1050
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   2
         Left            =   1815
         MaxLength       =   10
         TabIndex        =   0
         Tag             =   "Código Postal|T|S|||clientes|codposta|||"
         Top             =   930
         Width           =   1050
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   4905
         TabIndex        =   3
         Top             =   2550
         Width           =   975
      End
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   3720
         TabIndex        =   2
         Top             =   2550
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "Fecha "
         ForeColor       =   &H00972E0B&
         Height          =   255
         Index           =   16
         Left            =   570
         TabIndex        =   7
         Top             =   690
         Width           =   1815
      End
      Begin VB.Label Label4 
         Caption         =   "Desde"
         Height          =   195
         Index           =   15
         Left            =   930
         TabIndex        =   6
         Top             =   930
         Width           =   465
      End
      Begin VB.Label Label4 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   14
         Left            =   930
         TabIndex        =   5
         Top             =   1290
         Width           =   420
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   2
         Left            =   1500
         Picture         =   "frmTraspasoTPV.frx":000C
         ToolTipText     =   "Buscar fecha"
         Top             =   930
         Width           =   240
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   3
         Left            =   1500
         Picture         =   "frmTraspasoTPV.frx":0097
         ToolTipText     =   "Buscar fecha"
         Top             =   1290
         Width           =   240
      End
   End
End
Attribute VB_Name = "frmTraspasoTPV"
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

Private WithEvents frmC As frmCal 'calendario fechas
Attribute frmC.VB_VarHelpID = -1


Dim indCodigo As Integer 'indice para txtCodigo
Dim indFrame As Single 'nº de frame en el que estamos
 
Dim PrimeraVez As Boolean

Private Sub KEYpress(KeyAscii As Integer)
    If KeyAscii = 13 Then 'ENTER
        KeyAscii = 0
        SendKeys "{tab}"
    ElseIf KeyAscii = 27 Then Unload Me  'ESC
    End If
End Sub

Private Sub cmdAceptar_Click()
Dim cDesde As String, cHasta As String 'cadena codigo Desde/Hasta
Dim nDesde As String, nHasta As String 'cadena Descripcion Desde/Hasta
Dim cadTABLA As String, cOrden As String
Dim i As Byte

    
    If MsgBox("Desea continuar", vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
        MsgBox "meteremos la funcion"
    Else
    
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

 
    '###Descomentar
'    CommitConexion
         
    FrameCobrosVisible True, h, w
    indFrame = 5
            
    'Esto se consigue poneinedo el cancel en el opcion k corresponda
    Me.cmdCancel.Cancel = True
'    Me.Width = w + 70
'    Me.Height = h + 350
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
    KEYpress KeyAscii
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

Private Sub CargarTablaTemporal(DesFec As String, HasFec As String, SoloDesc As Byte)
    Dim SQL As String
    Dim Sql1 As String
    Dim Sql2 As String
    Dim sql3 As String
    Dim RS As ADODB.Recordset
    Dim Rs1 As ADODB.Recordset
    Dim Rs2 As ADODB.Recordset

    On Error GoTo eCargarTablaTemporal

    ' primero borramos los registros del usuario
    SQL = "delete from tmpinformes where codusu = " & vSesion.Empleado
    conn.Execute SQL
    

    ' cargamos la tabla temporal para el listado agrupando por fecha y turno
    ' unicamente cargamos el importe de mangueras, el resto lo inicializamos a 0
    SQL = "select fechatur, codturno, sum(importel) from sturno where "
    If DesFec <> "" Then
        SQL = SQL & " fechatur >= '" & Format(DesFec, FormatoFecha) & "'"
    End If
    If HasFec <> "" Then
        SQL = SQL & " and fechatur <= '" & Format(HasFec, FormatoFecha) & "'"
    End If
    SQL = SQL & " and tipocred = 0 "
    SQL = SQL & " group by fechatur, codturno "
    SQL = SQL & " order by fechatur, codturno "
    
    Set RS = New ADODB.Recordset ' Crear objeto
    RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText ' abrir cursor
      
    If Not RS.EOF Then RS.MoveFirst
    
    While Not RS.EOF
        Sql1 = "insert into tmpinformes (codusu, fecha1, campo1, importe1, importe2, importe3, importe4) "
        Sql1 = Sql1 & "values (" & vSesion.Empleado & ",'" & Format(RS.Fields(0).Value, FormatoFecha) & "',"
        If Not IsNull(RS.Fields(2).Value) Then
            Sql1 = Sql1 & RS.Fields(1).Value & "," & TransformaComasPuntos(ImporteSinFormato(RS.Fields(2).Value)) & ",0,0,0)"
        Else
            Sql1 = Sql1 & RS.Fields(1).Value & ",0,0,0,0)"
        End If
        
        conn.Execute Sql1
    
        RS.MoveNext
    Wend
    
    RS.Close
    SQL = "select fecha1, campo1 from tmpinformes where codusu = " & vSesion.Empleado
    SQL = SQL & " order by 1, 2"
    
    RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText ' abrir cursor
    If Not RS.EOF Then RS.MoveFirst
    While Not RS.EOF
        Sql1 = "select sum(importel) from scaalb where fecalbar = '" & Format(RS.Fields(0).Value, FormatoFecha)
        Sql1 = Sql1 & "' and codturno = " & RS.Fields(1).Value & " and codartic >=1 and codartic <= 9 "
    
        Set Rs1 = New ADODB.Recordset
        Rs1.Open Sql1, conn, adOpenForwardOnly, adLockPessimistic, adCmdText ' abrir cursor
        If Not Rs1.EOF Then Rs1.MoveFirst
        
        Sql2 = "select sum(importel) from scaalb where fecalbar = '" & Format(RS.Fields(0).Value, FormatoFecha)
        Sql2 = Sql2 & "' and codturno = " & RS.Fields(1).Value & " and codartic >=1 and codartic <= 9 "
        Sql2 = Sql2 & " and numalbar = 'MANUAL'"
        
        Set Rs2 = New ADODB.Recordset
        Rs2.Open Sql2, conn, adOpenForwardOnly, adLockPessimistic, adCmdText ' abrir cursor
        If Not Rs2.EOF Then Rs2.MoveFirst
        
        
        sql3 = "update tmpinformes set importe2 = "
        If Not IsNull(Rs1.Fields(0).Value) Then
            sql3 = sql3 & TransformaComasPuntos(ImporteSinFormato(Rs1.Fields(0).Value)) & ", "
        Else
            sql3 = sql3 & "0,"
        End If
        
        If Not IsNull(Rs2.Fields(0).Value) Then
            sql3 = sql3 & "importe4 = " & TransformaComasPuntos(ImporteSinFormato(Rs2.Fields(0).Value))
        Else
            sql3 = sql3 & "importe4 = 0 "
        End If
        
        sql3 = sql3 & " where fecha1 = '" & Format(RS.Fields(0).Value, FormatoFecha) & "' and "
        sql3 = sql3 & " campo1 = " & RS.Fields(1).Value
        sql3 = sql3 & " and codusu = " & vSesion.Empleado
        
        
        conn.Execute sql3
        
        Set Rs1 = Nothing
        Set Rs2 = Nothing
        
        Debug.Print RS.Fields(0).Value & "-" & RS.Fields(1).Value
        
        RS.MoveNext
    Wend

    ' una vez cargada la tabla temporal acualizamos el importe3 = diferencia entre importe1 e importe2
    SQL = "update tmpinformes set importe3 = importe1 - importe2 where codusu = " & vSesion.Empleado
    conn.Execute SQL

    If SoloDesc = 1 Then
        SQL = "delete from tmpinformes where codusu = " & vSesion.Empleado
        SQL = SQL & " and importe3 > -1 and importe3 < 1 "
        
        conn.Execute SQL
    End If

eCargarTablaTemporal:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Error en la carga de la tabla temporal"
    End If
End Sub
