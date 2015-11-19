VERSION 5.00
Begin VB.Form frmIdentifica 
   BackColor       =   &H00800000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   5745
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7650
   LinkTopic       =   "Form1"
   ScaleHeight     =   5745
   ScaleWidth      =   7650
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      IMEMode         =   3  'DISABLE
      Index           =   1
      Left            =   4320
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   4920
      Width           =   3015
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   0
      Left            =   4320
      TabIndex        =   0
      Top             =   3960
      Width           =   3015
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   195
      Left            =   2160
      TabIndex        =   5
      Top             =   5430
      Width           =   1725
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Cargando ....."
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
      Height          =   375
      Index           =   2
      Left            =   5040
      TabIndex        =   4
      Top             =   4920
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Usuario"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Index           =   0
      Left            =   4320
      TabIndex        =   2
      Top             =   3600
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Index           =   1
      Left            =   4320
      TabIndex        =   3
      Top             =   4560
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Image Image1 
      Height          =   5748
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   7692
   End
End
Attribute VB_Name = "frmIdentifica"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' +-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-
' +-+- Autor: DAVID (refet per CÈSAR) +-+-
' +-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-

Option Explicit




Dim PrimeraVez As Boolean
Dim T1 As Single

Private Sub Form_Activate()
    If PrimeraVez Then
        PrimeraVez = False

         Me.Refresh
         PonerVisible True
         If Text1(0).Text <> "" Then
            PonerFoco Text1(1)
         Else
            PonerFoco Text1(0)
         End If
             
         'Leemos el ultimo usuario conectado
         NumeroEmpresaMemorizar True
         
         T1 = T1 + 2.5 - Timer
         If T1 > 0 Then espera T1

         
         PonerVisible True
         If Text1(0).Text <> "" Then
            Text1(1).SetFocus
        Else
            Text1(0).SetFocus
        End If

    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
'    Screen.MousePointer = vbHourglass
    'PonerVisible False
'    T1 = Timer
    'Text1(0).Text = "root"
 '   Text1(1).Text = "aritel"
    PrimeraVez = True
    CargaImagen
    Label2.Caption = "Ver. " & App.Major & "." & App.Minor & "." & App.Revision

End Sub

Private Sub CargaImagen()
    On Error Resume Next
    Me.Image1 = LoadPicture(App.path & "\entrada.dat")
    If Err.Number <> 0 Then
        MsgBox Err.Description & vbCrLf & vbCrLf & "Error cargando", vbCritical
        Set Conn = Nothing
        End
    End If
End Sub


Private Sub Form_Unload(Cancel As Integer)
    NumeroEmpresaMemorizar False
End Sub

Private Sub Text1_GotFocus(Index As Integer)
    With Text1(Index)
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, 0, cerrar
    If cerrar Then Unload Me
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub Text1_LostFocus(Index As Integer)
    Text1(Index).Text = Trim(Text1(Index).Text)

    'Comprobamos si los dos estan con datos
    If Text1(0).Text <> "" And Text1(1).Text <> "" Then
        Validar
    End If
End Sub


Private Sub Validar()
Dim Ok As Byte
Dim cad As String
Dim SQL As String



    Set vSesion = New CSesion

    If vSesion.Leer(Text1(0).Text) = 0 Then
        'Con exito
        If vSesion.PasswdPROPIO = Text1(1).Text Then
            Ok = 0
        Else
            Ok = 1
        End If
    Else
        If Text1(0).Text = "root" And Text1(1).Text = "aritel" Then
            cad = "insert into usuarios (codusu, nomusu, login, passwordpropio, nivelusuges) "
            cad = cad & " values (0,'root','root','aritel',0)"
            Conn.Execute cad
            Ok = 0
        Else
            Ok = 2
        End If

    End If

    If Ok <> 0 Then
        MsgBox "Usuario o Password Incorrecto", vbExclamation

        Text1(1).Text = ""
        PonerFoco Text1(0)
    Else
        'OK
        If vSesion.Nivel < 0 Then
            MsgBox "Usuario sin Permisos.", vbExclamation
            End
        Else
            PonerVisible False
            Me.Refresh
            espera 0.2
        
            '[Monica]12/03/2015: para el caso de Alzira es multiempresa
            SQL = "select count(*) from usuarios.empresasarigasol "
            If TotalRegistros(SQL) > 1 Then
                CadenaDesdeOtroForm = "OK"
            Else
                CadenaDesdeOtroForm = "1"
            End If


            Unload Me
            
        End If
    
    End If


End Sub


Private Sub PonerVisible(visible As Boolean)
    'Label1(2).visible = Not visible  'Cargando
    Text1(0).visible = visible
    Text1(1).visible = visible
    Label1(0).visible = visible
    Label1(1).visible = visible
End Sub

'Lo que haremos aqui es ver, o guardar, el ultimo numero de empresa
'a la que ha entrado, y el usuario
Private Sub NumeroEmpresaMemorizar(Leer As Boolean)
Dim nf As Integer
Dim cad As String
On Error GoTo ENumeroEmpresaMemorizar

    cad = App.path & "\ultusu.dat"
    If Leer Then
        If Dir(cad) <> "" Then
            nf = FreeFile
            Open cad For Input As #nf
            Line Input #nf, cad
            Close #nf
            cad = Trim(cad)
            
                'El primer pipe es el usuario
                Text1(0).Text = cad
    
        End If
    Else 'Escribir
        nf = FreeFile
        Open cad For Output As #nf
        cad = Text1(0).Text
        Print #nf, cad
        Close #nf
    End If
ENumeroEmpresaMemorizar:
    Err.Clear
End Sub

