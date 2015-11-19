VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmContCieTurno 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Integración Cierre Contable de Turno"
   ClientHeight    =   5730
   ClientLeft      =   45
   ClientTop       =   2430
   ClientWidth     =   6825
   Icon            =   "frmContCieTurno.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5730
   ScaleWidth      =   6825
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
      Height          =   5505
      Left            =   150
      TabIndex        =   10
      Top             =   120
      Width           =   6555
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
         Height          =   2295
         Left            =   540
         TabIndex        =   14
         Top             =   1590
         Width           =   5625
         Begin VB.TextBox txtcodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   7
            Left            =   2010
            MaxLength       =   10
            TabIndex        =   7
            Tag             =   "Cta.Dif.negativas|T|S|||sparam|ctanegtat|||"
            Top             =   1920
            Width           =   585
         End
         Begin VB.TextBox txtNombre 
            BackColor       =   &H80000018&
            Enabled         =   0   'False
            Height          =   285
            Index           =   7
            Left            =   2700
            Locked          =   -1  'True
            TabIndex        =   27
            Top             =   1920
            Width           =   2805
         End
         Begin VB.TextBox txtNombre 
            BackColor       =   &H80000018&
            Enabled         =   0   'False
            Height          =   285
            Index           =   2
            Left            =   2700
            Locked          =   -1  'True
            TabIndex        =   20
            Top             =   390
            Width           =   2775
         End
         Begin VB.TextBox txtcodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   2
            Left            =   2010
            MaxLength       =   10
            TabIndex        =   3
            Tag             =   "Cta.Contable|T|S|||sparam|ctaconta|||"
            Top             =   390
            Width           =   585
         End
         Begin VB.TextBox txtNombre 
            BackColor       =   &H80000018&
            Enabled         =   0   'False
            Height          =   285
            Index           =   4
            Left            =   2700
            Locked          =   -1  'True
            TabIndex        =   17
            Top             =   1140
            Width           =   2775
         End
         Begin VB.TextBox txtcodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   4
            Left            =   2010
            MaxLength       =   10
            TabIndex        =   5
            Tag             =   "Cta.Contable|T|S|||sparam|ctaconta|||"
            Top             =   1140
            Width           =   585
         End
         Begin VB.TextBox txtNombre 
            BackColor       =   &H80000018&
            Enabled         =   0   'False
            Height          =   285
            Index           =   5
            Left            =   2700
            Locked          =   -1  'True
            TabIndex        =   16
            Top             =   1530
            Width           =   2805
         End
         Begin VB.TextBox txtcodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   5
            Left            =   2010
            MaxLength       =   10
            TabIndex        =   6
            Tag             =   "Cta.Dif.negativas|T|S|||sparam|ctanegtat|||"
            Top             =   1530
            Width           =   585
         End
         Begin VB.TextBox txtcodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   3
            Left            =   2010
            MaxLength       =   10
            TabIndex        =   4
            Tag             =   "Código Postal|T|S|||clientes|codposta|||"
            Top             =   780
            Width           =   1050
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   7
            Left            =   1680
            ToolTipText     =   "Buscar Concepto"
            Top             =   1920
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Conc.Haber Resto"
            ForeColor       =   &H00972E0B&
            Height          =   195
            Index           =   1
            Left            =   150
            TabIndex        =   28
            Top             =   1920
            Width           =   1410
         End
         Begin VB.Image imgAyuda 
            Height          =   240
            Index           =   0
            Left            =   3150
            MousePointer    =   4  'Icon
            Tag             =   "-1"
            ToolTipText     =   "Ayuda"
            Top             =   810
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Número Diario "
            ForeColor       =   &H00972E0B&
            Height          =   195
            Index           =   0
            Left            =   150
            TabIndex        =   21
            Top             =   450
            Width           =   1350
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   2
            Left            =   1680
            ToolTipText     =   "Buscar Diario"
            Top             =   420
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Conc.Haber Efectivo"
            ForeColor       =   &H00972E0B&
            Height          =   195
            Index           =   25
            Left            =   150
            TabIndex        =   19
            Top             =   1530
            Width           =   1500
         End
         Begin VB.Label Label1 
            Caption         =   "Concepto al Debe"
            ForeColor       =   &H00972E0B&
            Height          =   195
            Index           =   24
            Left            =   150
            TabIndex        =   18
            Top             =   1170
            Width           =   1350
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   4
            Left            =   1680
            ToolTipText     =   "Buscar Concepto"
            Top             =   1170
            Width           =   240
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   5
            Left            =   1680
            ToolTipText     =   "Buscar Concepto"
            Top             =   1530
            Width           =   240
         End
         Begin VB.Image imgFec 
            Height          =   240
            Index           =   3
            Left            =   1680
            Picture         =   "frmContCieTurno.frx":000C
            ToolTipText     =   "Buscar fecha"
            Top             =   810
            Width           =   240
         End
         Begin VB.Label Label4 
            Caption         =   "Fecha"
            ForeColor       =   &H00972E0B&
            Height          =   255
            Index           =   0
            Left            =   150
            TabIndex        =   15
            Top             =   810
            Width           =   495
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Datos para Selección"
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
         Height          =   1365
         Left            =   540
         TabIndex        =   11
         Top             =   210
         Width           =   5595
         Begin VB.TextBox txtcodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   6
            Left            =   4140
            MaxLength       =   10
            TabIndex        =   1
            Tag             =   "Código Postal|T|S|||clientes|codposta|||"
            Top             =   510
            Width           =   1080
         End
         Begin VB.TextBox txtcodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   0
            Left            =   1590
            MaxLength       =   10
            TabIndex        =   0
            Tag             =   "Código Postal|T|S|||clientes|codposta|||"
            Top             =   510
            Width           =   1080
         End
         Begin VB.TextBox txtcodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   1
            Left            =   1590
            MaxLength       =   1
            TabIndex        =   2
            Top             =   930
            Width           =   330
         End
         Begin VB.Label Label2 
            Caption         =   "Hasta"
            Height          =   195
            Index           =   0
            Left            =   3240
            TabIndex        =   26
            Top             =   555
            Width           =   465
         End
         Begin VB.Image imgFec 
            Height          =   240
            Index           =   1
            Left            =   3840
            Picture         =   "frmContCieTurno.frx":0097
            ToolTipText     =   "Buscar fecha"
            Top             =   532
            Width           =   240
         End
         Begin VB.Label Label2 
            Caption         =   "Desde"
            Height          =   195
            Index           =   20
            Left            =   690
            TabIndex        =   25
            Top             =   555
            Width           =   465
         End
         Begin VB.Image imgFec 
            Height          =   240
            Index           =   0
            Left            =   1290
            Picture         =   "frmContCieTurno.frx":0122
            ToolTipText     =   "Buscar fecha"
            Top             =   532
            Width           =   240
         End
         Begin VB.Label Label4 
            Caption         =   "Fecha"
            ForeColor       =   &H00972E0B&
            Height          =   255
            Index           =   16
            Left            =   150
            TabIndex        =   13
            Top             =   330
            Width           =   1425
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Nº Turno"
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
            TabIndex        =   12
            Top             =   930
            Width           =   645
         End
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   4905
         TabIndex        =   9
         Top             =   4860
         Width           =   975
      End
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   3720
         TabIndex        =   8
         Top             =   4860
         Width           =   975
      End
      Begin MSComctlLib.ProgressBar Pb1 
         Height          =   285
         Left            =   690
         TabIndex        =   22
         Top             =   3930
         Width           =   5370
         _ExtentX        =   9472
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Label lblProgres 
         Height          =   285
         Index           =   0
         Left            =   720
         TabIndex        =   24
         Top             =   4200
         Width           =   5265
      End
      Begin VB.Label lblProgres 
         Height          =   285
         Index           =   1
         Left            =   690
         TabIndex        =   23
         Top             =   4560
         Width           =   5295
      End
   End
End
Attribute VB_Name = "frmContCieTurno"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' +-+-+-+-+-+-+-+-+-+-+-
' +-+- Autor MONICA +-+-
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
Private WithEvents frmConce As frmConceConta 'conceptos de contabilidad
Attribute frmConce.VB_VarHelpID = -1
Private WithEvents frmTDia As frmDiaConta 'diarios de contabilidad
Attribute frmTDia.VB_VarHelpID = -1


'GENERALES PARA PASARLE A CRYSTAL REPORT
Private cadFormula As String 'Cadena con la FormulaSelection para Crystal Report
Private cadParam As String 'Cadena con los parametros para Crystal Report
Private numParam As Byte 'Numero de parametros que se pasan a Crystal Report
Private cadSelect As String 'Cadena para comprobar si hay datos antes de abrir Informe
Private cadTitulo As String 'Titulo para la ventana frmImprimir
Private cadNombreRPT As String 'Nombre del informe

Dim indCodigo As Integer 'indice para txtCodigo
 
'Se inicializan para cada Informe (tabla de BD a la que hace referencia
Dim Codigo As String 'Código para FormulaSelection de Crystal Report
Dim TipCod As String

Dim PrimeraVez As Boolean

Dim FechasIguales As Boolean

Dim CadFechas As String


Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, 0, cerrar
    If cerrar Then Unload Me
End Sub

Private Sub cmdAceptar_Click()
Dim Sql As String
Dim I As Byte
Dim cadWHERE As String

    If Not DatosOk Then Exit Sub
             
    FechasIguales = (txtcodigo(0).Text = txtcodigo(6).Text)
    
    If FechasIguales Then
        Sql = "SELECT count(*)" & _
              " FROM srecau, sforpa " & _
              "WHERE srecau.fechatur = " & DBSet(txtcodigo(0).Text, "F") & " and " & _
                   " srecau.codturno = " & DBSet(txtcodigo(1).Text, "N") & " and " & _
                   " srecau.codforpa = sforpa.codforpa and " & _
                   " srecau.intconta = 0 and " & _
                   " sforpa.cuadresn = 1 and not sforpa.codmacta is null and mid(sforpa.codmacta,1,1) <> ' '"
    Else
        ' si han puesto mas de un dia, no se tiene en cuenta el turno, son todos los turnos entre esas fechas
        Sql = "SELECT count(*)" & _
              " FROM srecau, sforpa " & _
              "WHERE srecau.fechatur >= " & DBSet(txtcodigo(0).Text, "F") & " and " & _
                   " srecau.fechatur <= " & DBSet(txtcodigo(6).Text, "F") & " and " & _
                   " srecau.codforpa = sforpa.codforpa and " & _
                   " srecau.intconta = 0 and " & _
                   " sforpa.cuadresn = 1 and not sforpa.codmacta is null and mid(sforpa.codmacta,1,1) <> ' '"
                   
        CadFechas = DevuelveFechas(Sql)
        
    End If
    
    If RegistrosAListar(Sql) = 0 Then
        MsgBox "No existen datos a contabilizar a esa fecha.", vbExclamation
        ' añadido, se han de marcar como contabilizados
        If FechasIguales Then
            Sql = "update srecau set intconta = 1 where srecau.fechatur = " & DBSet(txtcodigo(0).Text, "F") & _
                 " and srecau.codturno = " & DBSet(txtcodigo(1).Text, "N")
        Else
'[Monica]19/12/2012: Corrijo error de si una fecha esta contabilizada que no la incluya en un rango de fechas mayor
'            Sql = "update srecau set intconta = 1 where srecau.fechatur >= " & DBSet(txtcodigo(0).Text, "F") & _
'                 " and srecau.fechatur <= " & DBSet(txtcodigo(6).Text, "F")
            Sql = "update srecau set intconta = 1 where srecau.fechatur in " & CadFechas

        End If
        Conn.Execute Sql
        Exit Sub
    End If
    
    
    If FechasIguales Then
        cadWHERE = " scaalb.fecalbar = " & DBSet(txtcodigo(0).Text, "F") & " and " & _
                   " scaalb.codturno = " & DBSet(txtcodigo(1).Text, "N") & " and " & _
                   " sforpa.contabilizasn = 1 "
    Else
'[Monica]19/12/2012: Corrijo error de si una fecha esta contabilizada que no la incluya en un rango de fechas mayor
'        cadWhere = " scaalb.fecalbar >= " & DBSet(txtcodigo(0).Text, "F") & " and " & _
'                   " scaalb.fecalbar <= " & DBSet(txtcodigo(6).Text, "F") & " and " & _
'                   " sforpa.contabilizasn = 1 "
        cadWHERE = " scaalb.fecalbar in " & CadFechas & " and " & _
                   " sforpa.contabilizasn = 1 "
    End If
    
    ContabilizarCierre (cadWHERE)
     'Eliminar la tabla TMP
    BorrarTMPErrComprob

    DesBloqueoManual ("CIEREC") 'CIErre RECaudacion
    
    
    
eError:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "No se ha realizado el proceso de contabilización de cierre de turno. Llame a soporte."
    End If
    
    Pb1.visible = False
    lblProgres(0).Caption = ""
    lblProgres(1).Caption = ""
    cmdCancel_Click
End Sub


Private Function DevuelveFechas(vSQL As String) As String
Dim Sql2 As String
Dim Rs As ADODB.Recordset
Dim CadResul As String

    On Error GoTo eDevuelveFechas

    CadResul = ""
    
    Sql2 = Replace(vSQL, "count(*)", "distinct fechatur")
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql2, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    While Not Rs.EOF
        CadResul = CadResul & DBSet(Rs!fechatur, "F") & ","
        Rs.MoveNext
    Wend

    If CadResul <> "" Then
        ' quitamos la ultima coma
        CadResul = Mid(CadResul, 1, Len(CadResul) - 1)
        CadResul = "(" & CadResul & ")"
    End If

    Set Rs = Nothing

    DevuelveFechas = CadResul

eDevuelveFechas:

End Function


Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    If PrimeraVez Then
        PrimeraVez = False
        PonerFoco txtcodigo(0)
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
Dim h As Integer, w As Integer
Dim List As Collection

    PrimeraVez = True
    limpiar Me

    'IMAGES para busqueda
     Me.imgBuscar(2).Picture = frmPpal.imgListImages16.ListImages(1).Picture
     Me.imgBuscar(4).Picture = frmPpal.imgListImages16.ListImages(1).Picture
     Me.imgBuscar(5).Picture = frmPpal.imgListImages16.ListImages(1).Picture
     Me.imgBuscar(7).Picture = frmPpal.imgListImages16.ListImages(1).Picture
     txtcodigo(0).Text = Format(Now, "dd/mm/yyyy")
     txtcodigo(6).Text = txtcodigo(0).Text
     
     txtcodigo(2).Text = Format(vParamAplic.NumDiario, "000")
     txtcodigo(4).Text = Format(vParamAplic.ConceptoDebe, "000")
     txtcodigo(5).Text = Format(vParamAplic.ConceptoHaber, "000")
     txtcodigo(7).Text = Format(vParamAplic.ConceptoHaberResto, "000")
     txtNombre(2).Text = DevuelveDesdeBDNew(cConta, "tiposdiario", "desdiari", "numdiari", txtcodigo(2).Text, "N")
     txtNombre(4).Text = PonerNombreConcepto(txtcodigo(4))
     txtNombre(5).Text = PonerNombreConcepto(txtcodigo(5))
     txtNombre(7).Text = PonerNombreConcepto(txtcodigo(7))
    '###Descomentar
'    CommitConexion
         
    FrameCobrosVisible True, h, w
    Pb1.visible = False
    
    imgAyuda(0).Picture = frmPpal.ImageListB.ListImages(10).Picture
    
    'Esto se consigue poneinedo el cancel en el opcion k corresponda
    Me.cmdCancel.Cancel = True
'   Me.Width = w + 70
'   Me.Height = h + 350
End Sub

Private Sub frmC_Selec(vFecha As Date)
 'Fecha
    txtcodigo(indCodigo) = Format(vFecha, "dd/mm/yyyy") 'CByte(imgFec(0).Tag)).Text = Format(vFecha, "dd/MM/yyyy")
End Sub

Private Sub frmTDia_DatoSeleccionado(CadenaSeleccion As String)
'Form de Consulta de Clientes
    txtcodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "00")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmConce_DatoSeleccionado(CadenaSeleccion As String)
'Form de Consulta de Clientes
    txtcodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub


Private Sub imgAyuda_Click(Index As Integer)
Dim vCadena As String
    Select Case Index
        Case 0
           ' "____________________________________________________________"
            vCadena = "En caso de que las fechas Desde y Hasta coincidan, la fecha del " & vbCrLf & _
                      "asiento generado será la de datos de Contabilización. " & vbCrLf & vbCrLf & _
                      "Si no coinciden las fechas Desde y Hasta, la fecha de dicho asiento" & vbCrLf & _
                      "será la fecha del turno que le corresponda." & vbCrLf
                      
                      
    End Select
    MsgBox vCadena, vbInformation, "Descripción de Ayuda"
    
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

    Select Case Index
        Case 0
            indCodigo = 0
        Case 1
            indCodigo = 6
    End Select

    ' ***canviar l'index de imgFec pel 1r index de les imagens de buscar data***
    imgFec(0).Tag = indCodigo 'independentment de les dates que tinga, sempre pose l'index en la 27
    If txtcodigo(indCodigo).Text <> "" Then frmC.NovaData = txtcodigo(indCodigo).Text

    frmC.Show vbModal
    Set frmC = Nothing
    PonerFoco txtcodigo(indCodigo) 'CByte(imgFec(0).Tag) + 1)
    ' ***************************
End Sub

Private Sub ListView1_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub imgBuscar_Click(Index As Integer)
   Select Case Index
        Case 2 ' TIPOS DE DIARIO
            AbrirFrmDiario (Index)
        
        Case 4, 5, 7 'CONCEPTOS CONTABLES
            AbrirFrmConceptos (Index)
        
    End Select
    PonerFoco txtcodigo(indCodigo)
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
'14/02/2007
'    KEYpress KeyAscii
' ahora
    If KeyAscii = teclaBuscar Then
        Select Case Index
            Case 2: KEYBusqueda KeyAscii, 2 'diario
            Case 4: KEYBusqueda KeyAscii, 4 'concepto al debe
            Case 5: KEYBusqueda KeyAscii, 5 'concepto al haber
            Case 7: KEYBusqueda KeyAscii, 7 'concepto al haber
            Case 0: KEYFecha KeyAscii, 0 'fecha de turno
            Case 3: KEYFecha KeyAscii, 3 'fecha de contabilizacion
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
    txtcodigo(Index).Text = Trim(txtcodigo(Index).Text)
    
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub

    
    Select Case Index
        Case 2 ' NUMERO DE DIARIO
            If txtcodigo(Index).Text <> "" Then
                txtNombre(Index).Text = ""
                txtNombre(Index).Text = DevuelveDesdeBDNew(cConta, "tiposdiario", "desdiari", "numdiari", txtcodigo(Index).Text, "N")
                If txtNombre(Index).Text = "" Then
                    MsgBox "Número de Diario no existe en la contabilidad. Reintroduzca.", vbExclamation
'                    PonerFoco txtcodigo(Index)
                End If
            End If
        
        Case 4, 5, 7 'CONCEPTOS
            If txtcodigo(Index).Text <> "" Then txtNombre(Index).Text = PonerNombreConcepto(txtcodigo(Index))
            If txtNombre(Index).Text = "" Then
                MsgBox "Número de Concepto no existe en la contabilidad. Reintroduzca.", vbExclamation
'                PonerFoco txtcodigo(Index)
            End If

        Case 0, 3, 6 'FECHAS
            If txtcodigo(Index).Text <> "" Then PonerFormatoFecha txtcodigo(Index)
            '26/03/2007 cuando cambien la fecha del cierre cambia la del asiento
            If Index = 0 Then 'And txtcodigo(3).Text = "" Then
                txtcodigo(3).Text = txtcodigo(0).Text
            End If
            
    End Select
End Sub

Private Sub FrameCobrosVisible(visible As Boolean, ByRef h As Integer, ByRef w As Integer)
    Me.FrameCobros.visible = visible
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

Private Sub AbrirFrmDiario(indice As Integer)
    indCodigo = indice
    Set frmTDia = New frmDiaConta
    frmTDia.DatosADevolverBusqueda = "0|1|"
    frmTDia.CodigoActual = txtcodigo(indCodigo)
    frmTDia.Show vbModal
    Set frmTDia = Nothing
End Sub

Private Sub AbrirFrmConceptos(indice As Integer)
    indCodigo = indice
    Set frmConce = New frmConceConta
    frmConce.DatosADevolverBusqueda = "0|1|"
    frmConce.CodigoActual = txtcodigo(indCodigo)
    frmConce.Show vbModal
    Set frmConce = Nothing
End Sub
 
Private Sub ContabilizarCierre(cadWHERE As String)
'Contabiliza Facturas de Clientes o de Proveedores
Dim Sql As String
Dim b As Boolean
Dim tmpErrores As Boolean 'Indica si se creo correctamente la tabla de errores
Dim CCoste As String
Dim cadTABLA As String

    Sql = "CIEREC" 'contabilizar CIERRE DE RECAUDACION

    'Bloquear para que nadie mas pueda contabilizar
    DesBloqueoManual (Sql)
    If Not BloqueoManual(Sql, "1") Then
        MsgBox "No se pueden Contabilizar Cierre de Recaudación. Hay otro usuario contabilizándolo.", vbExclamation
        Screen.MousePointer = vbDefault
        Exit Sub
    End If

    
    'Visualizar la barra de Progreso
    Me.Pb1.visible = True
'    Me.Pb1.Top = 3350
    
    
    '==========================================================
    'REALIZAR COMPROBACIONES ANTES DE CONTABILIZAR FACTURAS
    '==========================================================
    
    Me.lblProgres(0).Caption = "Comprobaciones: "
    CargarProgres Me.Pb1, 100
        
    ' nuevo
    b = CrearTMPErrComprob()
    If Not b Then Exit Sub
    
    'comprobar que todas las CUENTAS de las distintos clientes que vamos a
    'contabilizar existen en la Conta: ssocio.codmacta IN (conta.cuentas.codmacta)
    '-----------------------------------------------------------------------------
    Me.MousePointer = vbHourglass

    Me.lblProgres(1).Caption = "Comprobando Cuentas Contables en contabilidad ..."
    cadTABLA = "ssocio"
    b = ComprobarCtaContable(cadTABLA, 1, cadWHERE)
    IncrementarProgres Me.Pb1, 20
    Me.Refresh
    If Not b Then
        frmMensaje.OpcionMensaje = 2
        frmMensaje.Show vbModal
        Me.MousePointer = vbDefault
        Exit Sub
    End If
    
    Me.MousePointer = vbDefault
    
    'comprobar que todas las CUENTAS de las formas de pago que vamos a
    'contabilizar existen en la Conta: sforpa.ctaventa IN (conta.cuentas.codmacta)
    '-----------------------------------------------------------------------------
    Me.lblProgres(1).Caption = "Comprobando Cuentas Ctbles de Pago en contabilidad ..."
    b = ComprobarCtaContable(cadTABLA, 5)
    IncrementarProgres Me.Pb1, 20
    Me.Refresh
    If Not b Then
        frmMensaje.OpcionMensaje = 2
        frmMensaje.Show vbModal
        Exit Sub
    End If
    
    'comprobar que todas las CUENTAS de diferencias positivas existen
    'en la Conta: sparam.ctaposit IN (conta.cuentas.codmacta)
    '-----------------------------------------------------------------------------
    Me.lblProgres(1).Caption = "Comprobando Cuenta Ctble Diferencias Positivas en contabilidad ..."
    b = ComprobarCtaContable(cadTABLA, 6)
    IncrementarProgres Me.Pb1, 20
    Me.Refresh
    If Not b Then
        frmMensaje.OpcionMensaje = 2
        frmMensaje.Show vbModal
        Exit Sub
    End If
    'comprobar que todas las CUENTAS de diferencias negativas existen
    'en la Conta: sparam.ctanegtat IN (conta.cuentas.codmacta)
    '-----------------------------------------------------------------------------
    Me.lblProgres(1).Caption = "Comprobando Cuenta Ctble Diferencias Negativas en contabilidad ..."
    b = ComprobarCtaContable(cadTABLA, 7)
    IncrementarProgres Me.Pb1, 20
    Me.Refresh
    If Not b Then
        frmMensaje.OpcionMensaje = 2
        frmMensaje.Show vbModal
        Exit Sub
    End If
     
    
    '===========================================================================
    'CONTABILIZAR CIERRE
    '===========================================================================
    Me.lblProgres(0).Caption = "Contabilizar Cierre: "
    CargarProgres Me.Pb1, 10
    Me.lblProgres(1).Caption = "Insertando Asiento en Contabilidad..."
    
    
    If FechasIguales Then
        cadWHERE = "fechatur = " & DBSet(txtcodigo(0).Text, "F") & " and codturno = " & DBSet(txtcodigo(1).Text, "N")
    Else
        '[Monica]19/12/2012: Corrijo error de si una fecha esta contabilizada que no la incluya en un rango de fechas mayor
'        cadWhere = "fechatur >= " & DBSet(txtcodigo(0).Text, "F") & " and fechatur <= " & DBSet(txtcodigo(6).Text, "F")
        cadWHERE = "fechatur in " & CadFechas
    End If
    
    b = PasarCierreAContab(cadWHERE)
    
    If Not b Then
        If tmpErrores Then
            'Cargar un listview con la tabla TEMP de Errores y mostrar
            'las facturas que fallaron
            frmMensaje.OpcionMensaje = 10
            frmMensaje.Show vbModal
        Else
            MsgBox "No pueden mostrarse los errores.", vbInformation
        End If
    Else
        MsgBox "El proceso ha finalizado correctamente.", vbInformation
    End If
    
End Sub

Private Function DatosOk() As Boolean
Dim b As Boolean
Dim Orden1 As String
Dim Orden2 As String
Dim FFin As Date
Dim Sql As String

   b = True

   If txtcodigo(0).Text = "" And b Then
        MsgBox "Introduzca la Fecha Desde de recaudación a contabilizar.", vbExclamation
        b = False
        PonerFoco txtcodigo(0)
    End If
    
   If txtcodigo(6).Text = "" And b Then
        MsgBox "Introduzca la Fecha Hasta de recaudación a contabilizar.", vbExclamation
        b = False
        PonerFoco txtcodigo(6)
    End If
    
    'si la fechadesde y fechahasta coinciden debemos introducir el turno obligatoriamente
    If b And txtcodigo(0).Text = txtcodigo(6).Text Then
        If txtcodigo(1).Text = "" Then
            MsgBox "Introduzca Nº del Turno a contabilizar.", vbExclamation
            b = False
            PonerFoco txtcodigo(1)
        End If
    End If
    ' comprobamos que han introducido los datos de la contabilidad
    ' +++NUMERO DE DIARIO+++
    If txtcodigo(2).Text = "" And b Then
        MsgBox "Introduzca Nº de diario a contabilizar.", vbExclamation
        b = False
        PonerFoco txtcodigo(2)
    End If
    
    ' +++FECHA DE ENTRADA+++
    If txtcodigo(3).Text = "" And b Then
        MsgBox "Introduzca la fecha de entrada del asiento.", vbExclamation
        b = False
        PonerFoco txtcodigo(3)
    Else
        ' comprobamos que la contabilizacion se encuentre en los ejercicios contables
         Orden1 = ""
         Orden1 = DevuelveDesdeBDNew(cConta, "parametros", "fechaini", "", "", "", "", "", "", "", "", "", "")
    
         Orden2 = ""
         Orden2 = DevuelveDesdeBDNew(cConta, "parametros", "fechafin", "", "", "", "", "", "", "", "", "", "")
         FFin = CDate(Orden2)
         If Not (CDate(Orden1) <= CDate(txtcodigo(3).Text) And CDate(txtcodigo(3).Text) < CDate(Day(FIni) & "/" & Month(FIni) & "/" & Year(FIni) + 2)) Then
            MsgBox "La Fecha de la contabilización no es del ejercicio actual ni del siguiente. Reintroduzca.", vbExclamation
            b = False
            PonerFoco txtcodigo(3)
         End If
    End If
    
    
    ' +++CONCEPTO AL DEBE+++
    If txtcodigo(4).Text = "" And b Then
        MsgBox "Introduzca el Concepto al Debe.", vbExclamation
        b = False
        PonerFoco txtcodigo(4)
    End If
    
    ' +++CONCEPTO AL HABER para efectivo+++
    If b Then
        If txtcodigo(5).Text = "" Then
            MsgBox "Introduzca el Concepto al Haber para efectivo.", vbExclamation
            b = False
            PonerFoco txtcodigo(5)
        Else
            ' comprobamos que el tipo de concepto de contabilidad para efectivo sea del tipo correspondiente
            Sql = DevuelveDesdeBDNew(cConta, "conceptos", "EsEfectivo340", "codconce", txtcodigo(5).Text, "N")
            If Sql = "0" Then
                MsgBox "El codigo de concepto ha de ser de Efectivo. Revise.", vbExclamation
                b = False
                PonerFoco txtcodigo(5)
            End If
        End If
    End If
    ' +++CONCEPTO AL HABER para el resto+++
    If b Then
        If txtcodigo(7).Text = "" Then
            MsgBox "Introduzca el Concepto al Haber para el resto.", vbExclamation
            b = False
            PonerFoco txtcodigo(7)
        Else
            ' comprobamos que el tipo de concepto de contabilidad no sea para efectivo
            Sql = DevuelveDesdeBDNew(cConta, "conceptos", "EsEfectivo340", "codconce", txtcodigo(7).Text, "N")
            If Sql = "1" Then
                MsgBox "El codigo de concepto no ha de ser de Efectivo. Revise.", vbExclamation
                b = False
                PonerFoco txtcodigo(7)
            End If
        End If
    End If
    DatosOk = b
End Function

Private Function PasarCierreAContab(cadWHERE As String) As Boolean
Dim Sql As String
Dim Rs As ADODB.Recordset
Dim Rs2 As ADODB.Recordset
Dim b As Boolean
Dim I As Integer
Dim numlinea As Integer
Dim Mc As CContadorContab
Dim numdocum As String
Dim ampliacion As String
Dim ampliaciond As String
Dim ampliacionh1 As String
Dim ampliacionh As String
Dim ImporteD As Currency
Dim ImporteH As Currency
Dim Diferencia As Currency
Dim Obs As String
Dim cadMen As String
Dim cad As String
Dim CtaDifer As String

Dim FechaAsiento As String

    On Error GoTo EPasarCie

    PasarCierreAContab = False
    
    b = True
    
    ConnConta.BeginTrans
    Conn.BeginTrans
    
    
    ' Vamos a hacer un asiento por cada fecha / turno
    Sql = "delete from tmpinformes where codusu = " & vSesion.Codigo
    Conn.Execute Sql
    
    ' insertamos los registros en la temporal
    Sql = "insert into tmpinformes (codusu, fecha1, codigo1) "
    Sql = Sql & " SELECT distinct " & vSesion.Codigo & ", fechatur, codturno " & _
              " FROM srecau, sforpa " & _
              "WHERE srecau.codforpa = sforpa.codforpa and " & _
                   " srecau.intconta = 0 and " & _
                   " sforpa.cuadresn = 1 and not sforpa.codmacta is null and mid(sforpa.codmacta,1,1) <> ' ' and " & cadWHERE
    Conn.Execute Sql
    
    
'[Monica]15/10/2013: introducia asientos sin nada en srecau todo a la cuenta de diferencias. Quito la insercion de estos albaranes
'                    habia añadido la ultima condicion pero lo quito todos
'    Sql = "insert into tmpinformes (codusu, fecha1, codigo1) "
'    Sql = Sql & "SELECT distinct " & vSesion.Codigo & ", fecalbar, codturno " & _
'              " FROM scaalb, sforpa, ssocio  " & _
'              "WHERE scaalb.codforpa = sforpa.codforpa and " & _
'                   " scaalb.codsocio = ssocio.codsocio and " & _
'                   " sforpa.contabilizasn = 1 and " & Replace(cadWHERE, "fechatur", "fecalbar") & _
'                   " and (scaalb.fecalbar, scaalb.codturno) in (select fecha1, codigo1 from tmpinformes where codusu = " & vSesion.Codigo & ")"
'    Conn.Execute Sql
    
    Sql = "select fecha1, codigo1 from tmpinformes where codusu = " & DBSet(vSesion.Codigo, "N")
    Sql = Sql & " group by 1, 2 "
    Sql = Sql & " order by 1, 2 "
    
    Set Rs2 = New ADODB.Recordset
    Rs2.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    While Not Rs2.EOF And b
        
        'Total de lineas de asiento a Insertar en la contabilidad
        Sql = "SELECT count(codmacta)" & _
              " FROM srecau, sforpa " & _
              "WHERE srecau.fechatur = " & DBSet(Rs2.Fields(0).Value, "F") & " and " & _
                   " srecau.codturno = " & DBSet(Rs2.Fields(1).Value, "N") & " and " & _
                   " srecau.codforpa = sforpa.codforpa and " & _
                   " srecau.intconta = 0 and " & _
                   " sforpa.cuadresn = 1 and not sforpa.codmacta is null and mid(sforpa.codmacta,1,1) <> ' '"
                 
        numlinea = TotalRegistros(Sql)
        
        Sql = "SELECT count(distinct ssocio.codmacta)" & _
              " FROM scaalb, sforpa, ssocio  " & _
              "WHERE scaalb.fecalbar = " & DBSet(Rs2.Fields(0).Value, "F") & " and " & _
                   " scaalb.codturno = " & DBSet(Rs2.Fields(1).Value, "N") & " and " & _
                   " scaalb.codforpa = sforpa.codforpa and " & _
                   " scaalb.codsocio = ssocio.codsocio and " & _
                   " sforpa.contabilizasn = 1 and " & _
                   " sforpa.tipforpa = 0 "
        numlinea = numlinea + TotalRegistros(Sql)
        
'[Monica]03/01/2012: añadida la condicion de las formas de pago de contado de las que no lo son
        Sql = "SELECT count(distinct ssocio.codmacta)" & _
              " FROM scaalb, sforpa, ssocio  " & _
              "WHERE scaalb.fecalbar = " & DBSet(Rs2.Fields(0).Value, "F") & " and " & _
                   " scaalb.codturno = " & DBSet(Rs2.Fields(1).Value, "N") & " and " & _
                   " scaalb.codforpa = sforpa.codforpa and " & _
                   " scaalb.codsocio = ssocio.codsocio and " & _
                   " sforpa.contabilizasn = 1 and " & _
                   " sforpa.tipforpa <> 0 "
        numlinea = numlinea + TotalRegistros(Sql)
        
        If numlinea = 0 Then Exit Function
        
        If numlinea > 0 Then
            numlinea = numlinea + 1
            
            CargarProgres Me.Pb1, numlinea
            
            
            Set Mc = New CContadorContab
            
            If FechasIguales Then
                FechaAsiento = txtcodigo(3).Text
            Else
                FechaAsiento = Format(DBLet(Rs2.Fields(0).Value, "F"), "dd/mm/yyyy")
            End If
                    
            If Mc.ConseguirContador("0", (CDate(FechaAsiento) <= CDate(FFin)), True) = 0 Then
            
            Obs = "Cierre Turno de fecha " & Format(Rs2.Fields(0).Value, "dd/mm/yyyy") & " y turno T-" & Format(Rs2.Fields(1).Value, "0")
    
        
            'Insertar en la conta Cabecera Asiento
            b = InsertarCabAsientoDia(txtcodigo(2).Text, Mc.Contador, FechaAsiento, Obs, cadMen)
            cadMen = "Insertando Cab. Asiento: " & cadMen
            
            If b Then
                Sql = "SELECT sforpa.codforpa, sforpa.codmacta, sum(importel)" & _
                      " FROM srecau, sforpa " & _
                      " WHERE srecau.fechatur = " & DBSet(Rs2.Fields(0).Value, "F") & " and " & _
                            " srecau.codturno = " & DBSet(Rs2.Fields(1).Value, "N") & " and " & _
                            " srecau.codforpa = sforpa.codforpa and " & _
                            " srecau.intconta = 0 and " & _
                            " sforpa.cuadresn = 1 and not sforpa.codmacta is null and mid(sforpa.codmacta,1,1) <> ' '" & _
                      " GROUP BY codforpa, codmacta "
                
                Set Rs = New ADODB.Recordset
                
                Rs.Open Sql, Conn, adOpenDynamic, adLockOptimistic, adCmdText
                
                I = 0
                ImporteD = 0
                ImporteH = 0
                
                numdocum = Format(DBLet(Rs2.Fields(0).Value, "F"), "ddmmyy") & "-T" & Format(DBLet(Rs2.Fields(1).Value, "N"), "0")
    '            ampliacion = "Cierre Turno " & Format(txtcodigo(0).Text, "dd/mm/yyyy") & " T-" & Format(txtcodigo(1).Text, "0")
                ampliacion = "CTu." & Format(DBLet(Rs2.Fields(0).Value, "F"), "dd/mm/yy") & "-T" & Format(DBLet(Rs2.Fields(1).Value, "N"), "0")
                ampliaciond = Trim(DevuelveDesdeBDNew(cConta, "conceptos", "nomconce", "codconce", txtcodigo(4).Text, "N")) & " " & ampliacion
                ampliacionh = Trim(DevuelveDesdeBDNew(cConta, "conceptos", "nomconce", "codconce", txtcodigo(5).Text, "N")) & " " & ampliacion
                '[Monica]15/10/2013: faltaba añadir el concepto del haber del resto
                ampliacionh1 = Trim(DevuelveDesdeBDNew(cConta, "conceptos", "nomconce", "codconce", txtcodigo(7).Text, "N")) & " " & ampliacion
                
                
                If Not Rs.EOF Then Rs.MoveFirst
                While Not Rs.EOF And b
                    I = I + 1
                    
                    cad = DBSet(txtcodigo(2).Text, "N") & "," & DBSet(FechaAsiento, "F") & "," & DBSet(Mc.Contador, "N") & ","
                    cad = cad & DBSet(I, "N") & "," & DBSet(Rs!Codmacta, "T") & "," & DBSet(numdocum, "T") & ","
                    
                    ' COMPROBAMOS EL SIGNO DEL IMPORTE PQ NO PERMITIMOS INTRODUCIR APUNTES CON IMPORTES NEGATIVOS
                    If Rs.Fields(2).Value > 0 Then
                        ' importe al debe en positivo
                        cad = cad & DBSet(txtcodigo(4).Text, "N") & "," & DBSet(ampliaciond, "T") & "," & DBSet(Rs.Fields(2).Value, "N") & ","
                        cad = cad & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & ",0"
                    
                        ImporteD = ImporteD + CCur(Rs.Fields(2).Value)
                    Else
                        ' importe al haber en positivo, cambiamos el signo
                        cad = cad & DBSet(txtcodigo(5).Text, "N") & "," & DBSet(ampliacionh, "T") & "," & ValorNulo & ","
                        cad = cad & DBSet((Rs.Fields(2).Value * -1), "N") & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & ",0"
                    
                        ImporteH = ImporteH + (CCur(Rs.Fields(2).Value) * (-1))
                    End If
                    
                    cad = "(" & cad & ")"
                    
                    b = InsertarLinAsientoDia(cad, cadMen)
                    cadMen = "Insertando Lin. Asiento: " & I
                
                    IncrementarProgres Me.Pb1, 1
                    Me.lblProgres(1).Caption = "Insertando línea de Asiento en Contabilidad...   (" & I & " de " & numlinea & ")"
                    Me.Refresh
                
                    Rs.MoveNext
                Wend
                Rs.Close
                
                If b Then
                    Sql = "SELECT ssocio.codmacta, 0 as tipo, sum(importel) " & _
                          "FROM scaalb, ssocio, sforpa " & _
                          "WHERE scaalb.fecalbar = " & DBSet(Rs2.Fields(0).Value, "F") & " and " & _
                          " scaalb.codturno = " & DBSet(Rs2.Fields(1).Value, "N") & " and " & _
                          " scaalb.codforpa = sforpa.codforpa and " & _
                          " scaalb.codsocio = ssocio.codsocio and " & _
                          " sforpa.contabilizasn = 1 and " & _
                          " sforpa.tipforpa = 0 " & _
                          " GROUP BY ssocio.codmacta, 2 " & _
                          " UNION " & _
                          "SELECT ssocio.codmacta, 1 as tipo , sum(importel) " & _
                          "FROM scaalb, ssocio, sforpa " & _
                          "WHERE scaalb.fecalbar = " & DBSet(Rs2.Fields(0).Value, "F") & " and " & _
                          " scaalb.codturno = " & DBSet(Rs2.Fields(1).Value, "N") & " and " & _
                          " scaalb.codforpa = sforpa.codforpa and " & _
                          " scaalb.codsocio = ssocio.codsocio and " & _
                          " sforpa.contabilizasn = 1 and " & _
                          " sforpa.tipforpa <> 0 " & _
                          " GROUP BY ssocio.codmacta, 2 " & _
                          " ORDER BY 1,2 "
    
                    Rs.Open Sql, Conn, adOpenDynamic, adLockOptimistic, adCmdText
                    
                    If Not Rs.EOF Then Rs.MoveFirst
                    
                    While Not Rs.EOF And b
                        I = I + 1
                    
                        cad = DBSet(txtcodigo(2).Text, "N") & "," & DBSet(FechaAsiento, "F") & "," & DBSet(Mc.Contador, "N") & ","
                        cad = cad & DBSet(I, "N") & "," & DBSet(Rs!Codmacta, "T") & "," & DBSet(numdocum, "T") & ","
                    
                        ' COMPROBAMOS EL SIGNO DEL IMPORTE SI ES NEGATIVO LO PONEMOS EN EL DEBE CON SIGNO POSITIVO
                        If Rs.Fields(2).Value > 0 Then
                            '[Monica]03/01/2013: si es efectivo el concepto al debe es el de efectivo
                            If Rs.Fields(1).Value = 0 Then
                                cad = cad & DBSet(txtcodigo(5).Text, "N") & "," & DBSet(ampliacionh, "T") & "," & ValorNulo & "," & DBSet(Rs.Fields(2).Value, "N") & ","
                                cad = cad & ValorNulo & "," & ValorNulo & "," & ValorNulo & ",0"
                            Else
                                cad = cad & DBSet(txtcodigo(7).Text, "N") & "," & DBSet(ampliacionh1, "T") & "," & ValorNulo & "," & DBSet(Rs.Fields(2).Value, "N") & ","
                                cad = cad & ValorNulo & "," & ValorNulo & "," & ValorNulo & ",0"
                            End If
                            ImporteH = ImporteH + CCur(Rs.Fields(2).Value)
                        
                        Else
                            cad = cad & DBSet(txtcodigo(4).Text, "N") & "," & DBSet(ampliaciond, "T") & "," & DBSet((Rs.Fields(2).Value * -1), "N") & "," & ValorNulo & ","
                            cad = cad & ValorNulo & "," & ValorNulo & "," & ValorNulo & ",0"
                        
                            ImporteD = ImporteD + (CCur(Rs.Fields(2).Value) * (-1))
                        
                        End If
                        cad = "(" & cad & ")"
                        
                        b = InsertarLinAsientoDia(cad, cadMen)
                        cadMen = "Insertando Lin. Asiento: " & I
                    
                        IncrementarProgres Me.Pb1, 1
                        Me.lblProgres(1).Caption = "Insertando línea de Asiento en Contabilidad...   (" & I & " de " & numlinea & ")"
                        Me.Refresh
                    
                        Rs.MoveNext
                    Wend
                    Rs.Close
                
                
                    If b Then
                        ' insertamos una linea al haber con la diferencia
                        If ImporteD <> ImporteH Then
                            I = I + 1
                            
                            If ImporteD > ImporteH Then
                                Diferencia = ImporteD - ImporteH
                                CtaDifer = vParamAplic.CtaPositiva
                            Else
                                Diferencia = ImporteH - ImporteD
                                CtaDifer = vParamAplic.CtaNegativa
                            End If
                            
                            cad = DBSet(txtcodigo(2).Text, "N") & "," & DBSet(FechaAsiento, "F") & "," & DBSet(Mc.Contador, "N") & ","
                            cad = cad & DBSet(I, "N") & "," & DBSet(CtaDifer, "T") & "," & DBSet(numdocum, "T") & ","
                            
                            If ImporteD < ImporteH Then
                                cad = cad & DBSet(txtcodigo(4).Text, "N") & "," & DBSet(ampliaciond, "T") & ","
                                cad = cad & DBSet(Diferencia, "N") & "," & ValorNulo & ","
                            Else
                                cad = cad & DBSet(txtcodigo(5).Text, "N") & "," & DBSet(ampliacionh, "T") & ","
                                cad = cad & ValorNulo & "," & DBSet(Diferencia, "N") & ","
                            End If
                            
                            cad = cad & ValorNulo & "," & ValorNulo & "," & ValorNulo & ",0"
                            
                            cad = "(" & cad & ")"
                        
                            b = InsertarLinAsientoDia(cad, cadMen)
                            cadMen = "Insertando Lin. Asiento: " & I
                
                            IncrementarProgres Me.Pb1, 1
                            Me.lblProgres(1).Caption = "Insertando línea de Asiento en Contabilidad...   (" & I & " de " & numlinea & ")"
                            Me.Refresh
                        End If
                        
                    End If
    
    ' de momento comentado para hacer pruebas
                    If b Then
                        'Poner intconta=1 en arigasol.srecau
                        b = ActualizarRecaudacion("fechatur = " & DBSet(Rs2.Fields(0).Value, "F") & " and codturno = " & DBSet(Rs2.Fields(1).Value, "N"), cadMen)
                        cadMen = "Actualizando Recaudación: " & cadMen
                    End If
                
            End If
        End If
       End If
       End If
       
       Rs2.MoveNext
       Set Mc = Nothing
       
    Wend
    
    Set Rs2 = Nothing
    
    '[Monica]04/10/2011: los turnos terceros de Alzira no se marcaban como contabilizados (sus formas de pago no se contabilizan)
    '                    pero hay que marcarlos para que no dé el aviso de que hay turnos pendientes de contabilizacion en la facturacion
    If b Then
        If Not FechasIguales Then
'[Monica]19/12/2012: Corrijo error de si una fecha esta contabilizada que no la incluya en un rango de fechas mayor
'           Sql = "update srecau set intconta = 1 where srecau.fechatur >= " & DBSet(txtcodigo(0).Text, "F") & _
'                 " and srecau.fechatur <= " & DBSet(txtcodigo(6).Text, "F")
           Sql = "update srecau set intconta = 1 where srecau.fechatur in " & CadFechas
           Conn.Execute Sql
        End If
    End If
    '04/10/2011
    
EPasarCie:
    If Err.Number <> 0 Then
        b = False
        MuestraError Err.Number, "Integrando Asiento a Contabilidad", Err.Description
    End If
    If b Then
        ConnConta.CommitTrans
        Conn.CommitTrans
        PasarCierreAContab = True
    Else
        ConnConta.RollbackTrans
        Conn.RollbackTrans
        PasarCierreAContab = False
    End If
End Function
