VERSION 5.00
Begin VB.Form frmEMail 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Enviar E-MAIL"
   ClientHeight    =   4290
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5745
   Icon            =   "frmEMail.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4290
   ScaleWidth      =   5745
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkCopia 
      Caption         =   "Copia remitente"
      Height          =   255
      Left            =   960
      TabIndex        =   20
      Top             =   3870
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Salir"
      Height          =   375
      Index           =   0
      Left            =   4320
      TabIndex        =   19
      Top             =   3840
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Enviar"
      Height          =   375
      Left            =   2940
      TabIndex        =   18
      Top             =   3840
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   3615
      Index           =   0
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   5715
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   0
         Left            =   960
         TabIndex        =   0
         Text            =   "Text1"
         Top             =   240
         Width           =   4455
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   1
         Left            =   960
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   660
         Width           =   4455
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   2
         Left            =   960
         TabIndex        =   2
         Text            =   "Text1"
         Top             =   1080
         Width           =   4455
      End
      Begin VB.TextBox Text1 
         Height          =   2055
         Index           =   3
         Left            =   960
         MultiLine       =   -1  'True
         TabIndex        =   3
         Text            =   "frmEMail.frx":000C
         Top             =   1560
         Width           =   4455
      End
      Begin VB.Label Label1 
         Caption         =   "Para"
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   11
         Top             =   300
         Width           =   330
      End
      Begin VB.Label Label1 
         Caption         =   "E-Mail"
         Height          =   195
         Index           =   1
         Left            =   180
         TabIndex        =   10
         Top             =   720
         Width           =   675
      End
      Begin VB.Label Label1 
         Caption         =   "Asunto"
         Height          =   255
         Index           =   2
         Left            =   180
         TabIndex        =   9
         Top             =   1140
         Width           =   555
      End
      Begin VB.Image Image1 
         Height          =   240
         Left            =   600
         Picture         =   "frmEMail.frx":0012
         Top             =   300
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Mensaje"
         Height          =   255
         Index           =   3
         Left            =   180
         TabIndex        =   8
         Top             =   1560
         Width           =   615
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   3675
      Index           =   1
      Left            =   0
      TabIndex        =   7
      Top             =   60
      Width           =   5715
      Begin VB.TextBox Text3 
         Height          =   1695
         Left            =   180
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   5
         Text            =   "frmEMail.frx":059C
         Top             =   1800
         Width           =   5355
      End
      Begin VB.TextBox Text2 
         Height          =   315
         Left            =   3120
         TabIndex        =   4
         Text            =   "Text2"
         Top             =   1080
         Width           =   2415
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Otro"
         Height          =   255
         Index           =   2
         Left            =   2460
         TabIndex        =   15
         Top             =   1140
         Width           =   675
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Error"
         Height          =   255
         Index           =   1
         Left            =   1560
         TabIndex        =   14
         Top             =   1140
         Width           =   855
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Sugerencia"
         Height          =   255
         Index           =   0
         Left            =   180
         TabIndex        =   13
         Top             =   1140
         Value           =   -1  'True
         Width           =   1335
      End
      Begin VB.Label Label3 
         Caption         =   "Mensaje"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   180
         TabIndex        =   17
         Top             =   1500
         Width           =   795
      End
      Begin VB.Label Label3 
         Caption         =   "Asunto"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   16
         Top             =   840
         Width           =   795
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Enviar e-Mail Ariadna Software"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   360
         Left            =   180
         TabIndex        =   12
         Top             =   300
         Width           =   4305
      End
   End
   Begin VB.Image Image2 
      Height          =   480
      Left            =   360
      Picture         =   "frmEMail.frx":05A2
      Top             =   3780
      Width           =   480
   End
End
Attribute VB_Name = "frmEMail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' +-+-+-+-+-+-+-+-+-+-+-+-+-+-+-
' +-+-    Autor: DAVID      +-+-
' +-+- Alguns canvis: CÈSAR +-+-
' +-+-+-+-+-+-+-+-+-+-+-+-+-+-+-

Option Explicit
Public Opcion As Byte
    '0 - Envio del PDF
    '1 - Envio Mail desde menu soporte
    
Private WithEvents frmB As frmBuscaGrid
Attribute frmB.VB_VarHelpID = -1
Private DatosaDevolverBusqueda As String

Dim cad As String
Dim HaDevueltoDatos As Boolean

Private Sub Enviar()
    Dim imageContentID, success
    Dim mailman As ChilkatMailMan
    Dim Valores As String
    
    On Error GoTo GotException
    Set mailman = New ChilkatMailMan
    
    'Esta cadena es constante de la lincencia comprada a CHILKAT
    mailman.UnlockComponent "1AriadnaMAIL_BOVuuRWYpC9f"
    mailman.LogMailSentFilename = App.path & "\mailSent.log"
    
' 06/02/2007
' modificacion: he descomentado este grupo de instrucciones
    'Servidor smtp
    Valores = ObtenerValoresEnvioMail  'Empipado: smtphost,smtpuser, pass, diremail
    If Valores = "" Then
        MsgBox "Falta configurar en parametros la opcion de envio mail(servidor, usuario, clave)"
        Exit Sub
    End If
    mailman.Smtphost = RecuperaValor(Valores, 1) ' vParam.SmtpHOST
    mailman.SmtpUsername = RecuperaValor(Valores, 2) 'vParam.SmtpUser
    mailman.SmtpPassword = RecuperaValor(Valores, 3) 'vParam.SmtpPass
    
    mailman.SmtpAuthMethod = "LOGIN"

' he comentado este grupo de instrucciones que estaban a piñon
'    mailman.SmtpHost = "ariadna.myariadna.com"
'    mailman.SmtpUsername = "manolo"
'    mailman.SmtpPassword = "aritel020763"
    
    ' Create the email, add content, address it, and sent it.
    Dim email As ChilkatEmail
    Set email = New ChilkatEmail
    
    'Si es de SOPORTE
'    If Opcion = 1 Then
'         'Obtenemos la pagina web de los parametros
'        '====David
''        Cad = DevuelveDesdeBD("mailsoporte", "parametros", "fechaini", Format(vParam.fechaini, FormatoFecha), "F")
'        '====
'        cad = DevuelveDesdeBDnew(conAri, "sparam", "maiempre", "codempre", 1, "N")
'        If cad = "" Then
'            MsgBox "Falta configurar en parametros el mail de soporte", vbExclamation
'            Exit Sub
'        End If
'
'        If cad = "" Then GoTo GotException
'        email.AddTo "Soporte Contabilidad", cad
'        cad = "Soporte Ariconta. "
'        If Option1(0).Value Then cad = cad & Option1(0).Caption
'        If Option1(1).Value Then cad = cad & Option1(1).Caption
'        If Option1(2).Value Then cad = cad & "Otro: " & Text2.Text
'        email.Subject = cad
'
'        'Ahora en text1(3).text generaremos nuestro mensaje
'        cad = "Fecha: " & Format(Now, "dd/mm/yyyy") & vbCrLf
'        cad = cad & "Hora: " & Format(Now, "hh:mm") & vbCrLf
'        cad = cad & "Usuario: " & vUsu.Nombre & vbCrLf
'        cad = cad & "Nivel USU: " & vUsu.Nivel & vbCrLf
'        cad = cad & "Empresa: " & vEmpresa.nomEmpre & vbCrLf
'        cad = cad & "&nbsp;<hr>"
'        cad = cad & Text3.Text & vbCrLf & vbCrLf
'        Text1(3).Text = cad
'    Else
        'Envio de mensajes normal
        email.AddTo Text1(0).Text, Text1(1).Text
        email.Subject = Text1(2).Text
        If chkCopia.Value = 1 Then email.AddBcc "AriGasol: " & vEmpresa.nomEmpre, RecuperaValor(Valores, 4)
'    End If
    
    'El resto lo hacemos comun
    
    
    
' 06/02/2007 todo esto lo he comentado y he añadido la parte de abajo
'    'La imagen
'    imageContentID = email.AddRelatedContent(App.path & "\logo.jpg")
'
'
'    cad = "<!DOCTYPE HTML PUBLIC " & Chr(34) & "-//W3C//DTD HTML 4.0 Transitional//EN" & Chr(34) & ">"
'    cad = cad & "<HTML><HEAD><TITLE>Mensaje</TITLE></HEAD>"
'    cad = cad & "<TABLE BORDER=""0"" CELLSPACING=1 CELLPADDING=0 WIDTH=576>"
'    'Cuerpo del mensaje
'    cad = cad & "<TR><TD VALIGN=""TOP""><P>"
'    FijarTextoMensaje
'    cad = cad & "</P></TD></TR>"
'    cad = cad & "<TR><TD VALIGN=""TOP""><P><HR ALIGN=""LEFT"" SIZE=1></P>"
'    'La imagen
'    cad = cad & "<P ALIGN=""CENTER""><IMG SRC=" & Chr(34) & "cid:" & imageContentID & Chr(34) & "></P>"
'    cad = cad & "<P ALIGN=""CENTER""><FONT SIZE=2>Mensaje creado desde el programa AriGasol de "
'    cad = cad & "<A HREF=""http://www.ariadnasoftware.com/"">Ariadna&nbsp;"
'    cad = cad & "Software S.L.</A></P><P ALIGN=""CENTER""></P>"
'    cad = cad & "<P>Este correo electrónico y sus documentos adjuntos están dirigidos EXCLUSIVAMENTE a "
'    cad = cad & " los destinatarios especificados. La información contenida puede ser CONFIDENCIAL"
'    cad = cad & " y/o estar LEGALMENTE PROTEGIDA.</P>"
'    cad = cad & "<P>Si usted recibe este mensaje por ERROR, por favor comuníqueselo inmediatamente al"
'    cad = cad & " remitente y ELIMÍNELO ya que usted NO ESTA AUTORIZADO al uso, revelación, distribución, "
'    cad = cad & " impresión o copia de toda o alguna parte de la información contenida, gracias "
'    cad = cad & ".</FONT></P><P><HR ALIGN=""LEFT"" SIZE=1></TD>"
'    cad = cad & "</TR></TABLE></BODY></HTML>"
    
    
    cad = "<!DOCTYPE HTML PUBLIC " & Chr(34) & "-//W3C//DTD HTML 4.0 Transitional//EN" & Chr(34) & ">"
    cad = cad & "<HTML><HEAD><TITLE>Mensaje</TITLE></HEAD>"
    cad = cad & "<TABLE BORDER=""0"" CELLSPACING=1 CELLPADDING=0 WIDTH=576>"
    'Cuerpo del mensaje
    cad = cad & "<TR><TD VALIGN=""TOP""><P>"
    FijarTextoMensaje
    cad = cad & "</P></TD></TR>"
    cad = cad & "<TR><TD VALIGN=""TOP""><P><hr></P>"
    'La imagen
    'cad = cad & "<P ALIGN=""CENTER""><IMG SRC=" & Chr(34) & "cid:" & imageContentID & Chr(34) & "></P>"
    cad = cad & "<P ALIGN=""CENTER""><FONT SIZE=2>Mensaje creado desde el programa " & App.EXEName & " de "
    cad = cad & "<A HREF=""http://www.ariadnasoftware.com/"">Ariadna&nbsp;"
    cad = cad & "Software S.L.</A></P><P ALIGN=""CENTER""></P>"
    cad = cad & "<P>Este correo electrónico y sus documentos adjuntos estan dirigidos EXCLUSIVAMENTE a "
    cad = cad & " los destinatarios especificados. La información contenida puesde ser CONFIDENCIAL"
    cad = cad & " y/o estar LEGALMENTE PROTEGIDA.</P>"
    cad = cad & "<P>Si usted recibe este mensaje por ERROR, por favor comuníqueselo inmediatamente al"
    cad = cad & " remitente y ELIMINELO ya que usted NO ESTA AUTORIZADO al uso, revelación, distribución"
    cad = cad & " impresión o copia de toda o alguna parte de la información contenida, Gracias "
    cad = cad & ".</FONT></P><P><HR ALIGN=""LEFT"" SIZE=1></TD>"
    cad = cad & "</TR></TABLE></BODY></HTML>"
     
    
    
    email.SetHtmlBody (cad)
    
    
    
    email.AddPlainTextAlternativeBody "Programa e-mail NO soporta HTML. " & vbCrLf & Text1(3).Text
    email.From = RecuperaValor(Valores, 4) 'vParam.diremail
    'email.From = "manolo@myariadna.com"
    
    If Opcion = 0 Then
        'ADjunatmos el PDF
        email.AddFileAttachment App.path & "\docum.pdf"
    End If
        
    
    'email.SendEncrypted = 1
    success = mailman.SendEmail(email)
    If (success = 1) Then
        cad = "Mensaje enviado correctamente."
        MsgBox cad, vbInformation
        Command2(0).SetFocus
    Else
        cad = "Han ocurrido errores durante el envio.Compruebe el archivo log.xml para mas informacion"
        mailman.SaveXmlLog App.path & "\log.xml"
        MsgBox cad, vbExclamation
    End If
    
    
GotException:
    If Err.Number <> 0 Then MuestraError Err.Number, Err.Description
    Set email = Nothing
    Set mailman = Nothing

End Sub

Private Sub Command1_Click()
    If Not DatosOK Then Exit Sub
    Screen.MousePointer = vbHourglass
    Image2.visible = True
    Me.Refresh
    Enviar
    Image2.visible = False
    Screen.MousePointer = vbDefault
End Sub

Private Sub Command2_Click(Index As Integer)
    Unload Me
End Sub

Private Sub Form_Load()
    Image2.visible = False
    Limpiar Me
    Frame1(0).visible = (Opcion = 0)
    Frame1(1).visible = (Opcion = 1)
    If Opcion = 1 Then HabilitarText

    '###Descomentar
'    cad = DevuelveDesdeBD("smtpHost", "spara1", "fechaini", Format(vParam.fechaini, FormatoFecha), "F")
'    cad = DevuelveDesdeBDnew(conAri, "spara1", "smtphost", "codigo", "1", "N")
'    Me.Command1.Enabled = (cad <> "")
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Opcion = 0
End Sub

Private Sub frmB_Selecionado(CadenaDevuelta As String)
    Text1(0).Text = RecuperaValor(CadenaDevuelta, 1)
    Text1(1).Text = RecuperaValor(CadenaDevuelta, 2)
End Sub

Private Sub Image1_Click()
    MandaBusquedaPrevia "maisocio <> '' and maisocio is not null and maisocio <> ' '"
End Sub

'Private Sub frmC_DatoSeleccionado(CadenaSeleccion As String)
'
'    Screen.MousePointer = vbHourglass
'    Text1(0).Tag = RecuperaValor(CadenaSeleccion, 1)
'    Text1(0).Text = RecuperaValor(CadenaSeleccion, 3)
'    'Si regresa con datos tengo k devolveer desde la bd el campo e-mail
'    Text1(1).Text = RecuperaValor(CadenaSeleccion, 4)
''    cad = DevuelveDesdeBDNew(conAri, "sclien", "maiclie1", "codclien", Text1(0).Tag, "T")
''    Text1(1).Text = cad
'    Screen.MousePointer = vbDefault
'End Sub
'
'Private Sub Image1_Click()
'    Set frmC = New frmFacClientes
'    frmC.DatosADevolverBusqueda = "0|1"
''    frmC.ConfigurarBalances = 5  'NUEVO opcion
'    frmC.Show vbModal
'    Set frmC = Nothing
'    If Text1(0).Text <> "" Then PonerFoco Text1(2)
'End Sub

Private Sub Option1_Click(Index As Integer)
    HabilitarText
End Sub

Private Sub Text1_GotFocus(Index As Integer)
    ConseguirFoco Text1(Index), 3
'    With Text1(Index)
'        .SelStart = 0
'        .SelLength = Len(.Text)
'    End With
End Sub

Private Function DatosOK() As Boolean
Dim i As Integer

    DatosOK = False
    If Opcion = 0 Then
                'Pocas cosas a comprobar
                For i = 0 To 2
                    Text1(i).Text = Trim(Text1(i).Text)
                    If Text1(i).Text = "" Then
                        MsgBox "El campo: " & Label1(i).Caption & " no puede estar vacio.", vbExclamation
                        Exit Function
                    End If
                Next i
                
                'EL del mail tiene k tener la arroba @
                i = InStr(1, Text1(1).Text, "@")
                If i = 0 Then
                    MsgBox "Direccion e-mail erronea", vbExclamation
                    Exit Function
                End If
    Else
        Text2.Text = Trim(Text2.Text)
        'SOPORTE
        If Trim(Text3.Text) = "" Then
            MsgBox "El mensaje no puede ir en blanco", vbExclamation
            Exit Function
        End If
        If Option1(2).Value Then
            If Text2.Text = "" Then
                MsgBox "El campo 'OTRO asunto' no puede ir en blanco", vbExclamation
                Exit Function
            End If
        End If
    End If
      
    'Llegados aqui OK
    DatosOK = True
        
End Function


Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    If Index = 3 Then Exit Sub 'Si estamos en el de datos nos salimos
    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{tab}"
    End If
End Sub

'El procedimiento servira para ir buscando los vbcrlf y cambiarlos por </p><p>
Private Sub FijarTextoMensaje()
Dim i As Integer
Dim J As Integer

    J = 1
    Do
        i = InStr(J, Text1(3).Text, vbCrLf)
        If i > 0 Then
              cad = cad & Mid(Text1(3).Text, J, i - J) & "</P><P>"
        Else
            cad = cad & Mid(Text1(3).Text, J)
        End If
        J = i + 2
    Loop Until i = 0
End Sub

Private Sub HabilitarText()
    If Option1(2).Value Then
        Text2.Enabled = True
        Text2.BackColor = vbWhite
    Else
        Text2.Enabled = False
        Text2.BackColor = &H80000018
    End If
End Sub

Private Function RecuperarDatosEMAILAriadna() As Boolean
Dim NF As Integer

    RecuperarDatosEMAILAriadna = False
    NF = FreeFile
    Open App.path & "\soporte.dat" For Input As #NF
    Line Input #NF, cad
    Close #NF
    If cad <> "" Then RecuperarDatosEMAILAriadna = True
    
End Function

Private Function ObtenerValoresEnvioMail() As String
Dim miRsAux As ADODB.Recordset

    ObtenerValoresEnvioMail = ""
    Set miRsAux = New ADODB.Recordset
    cad = "Select diremail,SmtpHost, SmtpUser, SmtpPass  from sparam where"
    '####Descomentar
'    Cad = Cad & " fechaini='" & Format(vParam.fechaini, FormatoFecha) & "';"
    cad = cad & " codparam=1;"
    miRsAux.Open cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not miRsAux.EOF Then
        cad = DBLet(miRsAux!Smtphost)
        cad = cad & "|" & DBLet(miRsAux!SmtpUser)
        cad = cad & "|" & DBLet(miRsAux!Smtppass)
        cad = cad & "|" & DBLet(miRsAux!DireMail) & "|"
        ObtenerValoresEnvioMail = cad
    End If
    miRsAux.Close
    Set miRsAux = Nothing
End Function


Private Sub MandaBusquedaPrevia(CadB As String)
    Dim cad As String
        
    'Cridem al form
    ' **************** arreglar-ho per a vore lo que es desije ****************
    ' NOTA: el total d'amples de ParaGrid, ha de sumar 100
    cad = ""
    cad = cad & "Codigo|codsocio|N||10·"
    cad = cad & "Nombre|nomsocio|T||60·"
    cad = cad & "Mail|maisocio|T||30·"
    If cad <> "" Then
        Screen.MousePointer = vbHourglass
        Set frmB = New frmBuscaGrid
        frmB.vCampos = cad
        frmB.vtabla = "ssocio"
        frmB.vSQL = CadB
        HaDevueltoDatos = False
        frmB.vDevuelve = "1|2|" '*** els camps que volen que torne ***
        frmB.vTitulo = "Clientes" ' ***** repasa açò: títol de BuscaGrid *****
        frmB.vSelElem = 1

        frmB.Show vbModal
        Set frmB = Nothing
        'Si ha posat valors i tenim que es formulari de búsqueda llavors
        'tindrem que tancar el form llançant l'event
        If HaDevueltoDatos Then
'            If (Not Data1.Recordset.EOF) And DatosADevolverBusqueda <> "" Then
'                cmdRegresar_Click
        Else   'de ha retornat datos, es a decir NO ha retornat datos
            PonerFoco Text1(0)
        End If
    End If
End Sub

