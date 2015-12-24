VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmBorreTurno 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Borre Turno completo"
   ClientHeight    =   3810
   ClientLeft      =   45
   ClientTop       =   2430
   ClientWidth     =   5970
   Icon            =   "frmBorreTurno.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3810
   ScaleWidth      =   5970
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtCodigo 
      Height          =   285
      Index           =   2
      Left            =   1680
      MaxLength       =   15
      TabIndex        =   0
      Top             =   1560
      Width           =   1095
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   3360
      TabIndex        =   2
      Top             =   3135
      Width           =   975
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   4485
      TabIndex        =   3
      Top             =   3135
      Width           =   975
   End
   Begin VB.TextBox txtCodigo 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   0
      Left            =   1680
      MaxLength       =   4
      TabIndex        =   1
      Top             =   2040
      Width           =   375
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   1920
      Top             =   3000
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Image imgFec 
      Height          =   240
      Index           =   2
      Left            =   1320
      Picture         =   "frmBorreTurno.frx":000C
      ToolTipText     =   "Buscar fecha"
      Top             =   1560
      Width           =   240
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Seleccione Fecha y Turno a eliminar:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   6
      Left            =   720
      TabIndex        =   6
      Top             =   600
      Width           =   4425
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Fecha"
      ForeColor       =   &H00000080&
      Height          =   195
      Index           =   3
      Left            =   720
      TabIndex        =   5
      Top             =   1560
      Width           =   450
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Turno"
      ForeColor       =   &H00000080&
      Height          =   195
      Index           =   0
      Left            =   720
      TabIndex        =   4
      Top             =   2040
      Width           =   420
   End
End
Attribute VB_Name = "frmBorreTurno"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' +-+-+-+-+-+-+-+-+-+-+-
' +-+- Autor:MANOLO +-+-
' +-+-+-+-+-+-+-+-+-+-+-

Option Explicit

Private WithEvents frmC As frmCal 'Calendario de Fechas
Attribute frmC.VB_VarHelpID = -1

'GENERALES PARA PASARLE A CRYSTAL REPORT
Private cadFormula As String 'Cadena con la FormulaSelection para Crystal Report
Private cadSelect As String 'Cadena para comprobar si hay datos antes de abrir Informe
Private cadParam As String 'Cadena con los parametros para Crystal Report
Private numParam As Byte 'Numero de parametros que se pasan a Crystal Report

Dim indCodigo As Integer 'indice para txtCodigo
 
'Se inicializan para cada Informe (tabla de BD a la que hace referencia
Dim tabla As String
Dim Codigo As String 'Código para FormulaSelection de Crystal Report
Dim TipCod As String

Dim PrimeraVez As Boolean

Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, 0, cerrar
    If cerrar Then Unload Me
End Sub

Private Sub cmdAceptar_Click()
'Obtener la cadena SQL para eliminar los registros seleccionados
Dim cDesde As String 'Turno
Dim fDesde As String 'Fecha
Dim SQL As String

    If Not DatosOk Then Exit Sub


    InicializarVbles
    SQL = ""
    'Valores para Formula seleccion del informe
    cDesde = Trim(txtCodigo(0).Text)
    fDesde = Trim(txtCodigo(2).Text)
    
    SQL = tabla & ".fecalbar=" & "'" & Format(fDesde, FormatoFecha) & "'"
    
    '[Monica]24/12/2015: el borrado es del día completo
    If vParamAplic.Cooperativa <> 2 Then
        SQL = SQL & " AND " & tabla & ".codturno=" & cDesde
    End If
    
'    AnyadirAFormula cadFormula, sql
'    AnyadirAFormula cadSelect, sql
    
    
    'en cadSelect tenemos el valor correcto de la WHERE para borrar los registros
    EliminarSelTurno SQL
    
End Sub


Private Function DatosOk() As Boolean
Dim b As Boolean
Dim Datos As String
Dim SQL As String

    On Error GoTo EDatosOK

    DatosOk = False
    
    b = True
    
    If txtCodigo(2).Text = "" Then
        MsgBox "Debe introducir una fecha. Revise.", vbExclamation
        PonerFoco txtCodigo(2)
        b = False
    End If
         
    If vParamAplic.Cooperativa <> 2 Then
        If b And txtCodigo(0).Text = "" Then
            MsgBox "Debe introducir un turno. Revise.", vbExclamation
            PonerFoco txtCodigo(0)
            b = False
        End If
    End If
    
    '[Monica]21/10/2015: si hay cargas de gasoleo profesional declaradas no podemos borrar turno
    If b Then
        If txtCodigo(0).Text <> "" And txtCodigo(2).Text <> "" Then
            SQL = "select count(*) " & _
                    " from scaalb as a, starje as b, ssocio as c, sartic as d" & _
                    " where a.numtarje in (select numtarje from starje where tiptarje = 2)" & _
                    " and a.codartic in (select codartic from sartic where gp = 1)" & _
                    " and b.numtarje = a.numtarje" & _
                    " and c.codsocio = a.codsocio" & _
                    " and d.codartic = a.codartic" & _
                    " and a.declaradogp = 1" & _
                    " and a.fecalbar = '" & Format(txtCodigo(2).Text, "yyyy-mm-dd") & "'" & _
                    " and a.codturno = " & DBSet(txtCodigo(0).Text, "N")

            If TotalRegistros(SQL) <> 0 Then
                MsgBox "Hay cargas de Gasóleo Profesional en el turno que han sido declaradas. " & vbCrLf & vbCrLf & "No se permite borrar el turno.", vbExclamation
                PonerFoco txtCodigo(2)
                b = False
            End If
        End If
    End If
    
    DatosOk = b
EDatosOK:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Function




Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    If PrimeraVez Then PrimeraVez = False
    
    Screen.MousePointer = vbDefault
    PonerFoco txtCodigo(2)
End Sub

Private Sub Form_Load()

    PrimeraVez = True
    limpiar Me
    
    '###Descomentar
'    CommitConexion
    
    tabla = "scaalb"
    
    '[Monica]24/12/2015: en Regaixo se borra el dia completo
    Label1(0).visible = (vParamAplic.Cooperativa <> 2)
    txtCodigo(0).visible = (vParamAplic.Cooperativa <> 2)
    Label1(0).Enabled = (vParamAplic.Cooperativa <> 2)
    txtCodigo(0).Enabled = (vParamAplic.Cooperativa <> 2)
    
    
    'Esto se consigue poneinedo el cancel en el opcion k corresponda
    Me.cmdCancel.Cancel = True
End Sub

Private Sub frmC_Selec(vFecha As Date)
   txtCodigo(CByte(imgFec(2).Tag)).Text = Format(vFecha, "dd/MM/yyyy")
End Sub

Private Sub imgFec_Click(Index As Integer)
    'Calendario de Fechas
    Dim esq As Long
    Dim dalt As Long
    Dim menu As Long
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
    
    menu = Me.Height - Me.ScaleHeight 'ací tinc el heigth del menú i de la toolbar
   
    ' es desplega baix i cap a la dreta
    frmC.Left = esq + imgFec(Index).Parent.Left + 30
    frmC.Top = dalt + imgFec(Index).Parent.Top + imgFec(Index).Height + menu - 40
    
    ' es desplega dalt i cap a la esquerra
    'frmC.Left = esq + imgFec(Index).Parent.Left - frmC.Width + imgFec(Index).Width + 40
    'frmC.Top = dalt + imgFec(Index).Parent.Top + Toolbar1.Height - frmC.Height + 20
    'frmC.Top = dalt + imgFec(Index).Parent.Top - frmC.Height + menu - 25

    ' ***canviar l'index de imgFec pel 1r index de les imagens de buscar data***
    imgFec(2).Tag = Index 'independentment de les dates que tinga, sempre pose l'index en la 27
    If txtCodigo(Index).Text <> "" Then frmC.NovaData = txtCodigo(Index).Text
    
    frmC.Show vbModal
    Set frmC = Nothing
    PonerFoco txtCodigo(CByte(imgFec(2).Tag))
    ' **************************************************************************
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
            Case 2: KEYFecha KeyAscii, 2 'fecha
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

    'Quitar espacios en blanco por los lados
    txtCodigo(Index).Text = Trim(txtCodigo(Index).Text)
    
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub
    
    Select Case Index
        Case 0 'TURNO
            If txtCodigo(Index).Text <> "" Then PonerFormatoEntero txtCodigo(Index)

        Case 2 'FECHA
            If txtCodigo(Index).Text <> "" Then PonerFormatoFecha txtCodigo(Index)
    End Select
End Sub

Private Sub InicializarVbles()
    cadFormula = ""
    cadSelect = ""
    cadParam = ""
    numParam = 0
End Sub

Private Function EliminarSelTurno(cadW As String) As Boolean
'Eliminar Albaranes de Fecha y Turno: Tabla (scaalb)
'que cumplan los criterios seleccionados en la cadena WHERE cadW

Dim cad As String, SQL As String
Dim Rs As ADODB.Recordset
Dim todasElim As Boolean

    On Error GoTo EEliminar


    If vParamAplic.Cooperativa <> 2 Then
        cad = "Va a eliminar el Turno seleccionado." & vbCrLf
        cad = cad & vbCrLf & vbCrLf & "¿Desea Eliminarlo? "
    Else
        cad = "Va a eliminar el día seleccionado." & vbCrLf
        cad = cad & vbCrLf & vbCrLf & "¿Desea Eliminarlo? "
    End If
    
    If MsgBox(cad, vbQuestion + vbYesNoCancel) = vbYes Then     'Borramos
        'Hay que eliminar
        On Error GoTo EEliminar
        Screen.MousePointer = vbHourglass
        
        
        If EliminarTurno(txtCodigo(2).Text, txtCodigo(0).Text) Then
            If vParamAplic.Cooperativa <> 2 Then
                MsgBox "El turno seleccionado se eliminó correctamente.", vbInformation
            Else
                MsgBox "El día seleccionado se eliminó correctamente.", vbInformation
            End If
            Unload Me
        Else
            MsgBox "ATENCIÓN: Se ha producido un error al eliminar.", vbInformation
            Unload Me
        End If
    End If
    Screen.MousePointer = vbDefault
    
EEliminar:
    Screen.MousePointer = vbDefault
    If Err.Number > 0 Then MuestraError Err.Number, "Eliminar Turno", Err.Description
End Function

Private Function EliminarTurno(fecAlbar As String, codTurno As String) As Boolean
'Eliminar las lineas y la Cabecera de un Caja. Tablas: cajascab, cajaslin
Dim SQL As String
Dim b As Boolean

    On Error GoTo EEliminarTur
    EliminarTurno = False
    b = False
    SQL = " WHERE  fecalbar='" & Format(fecAlbar, FormatoFecha) & "'  "
    
    '[Monica]24/12/2015: el borrado es de todo el dia
    If vParamAplic.Cooperativa <> 2 Then
        SQL = SQL & " AND codturno=" & codTurno
    End If
    
    Conn.BeginTrans
    
    'Cabecera
    Conn.Execute "DELETE FROM scaalb " & SQL
    
    SQL = " WHERE  fechatur='" & Format(fecAlbar, FormatoFecha) & "' "
    
    '[Monica]24/12/2015: el borrado es de todo el dia
    If vParamAplic.Cooperativa <> 2 Then
        SQL = SQL & " AND codturno=" & codTurno
    End If
    
    ' añadido 20/03/2007 el tipo debe ser <= 2 (tanques, contadores y ventastipo)
    
    Conn.Execute "DELETE FROM sturno " & SQL & " and tiporegi <= 2 "
    Conn.Execute "DELETE FROM srecau " & SQL
    
    EliminarTurno = True
    b = True
    
EEliminarTur:
    
    b = Not (Err.Number <> 0)
    
    If Not b Then
        Conn.RollbackTrans
    Else
        Conn.CommitTrans
    End If
    EliminarTurno = b
End Function
