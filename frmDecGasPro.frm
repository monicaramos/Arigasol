VERSION 5.00
Begin VB.Form frmDecGasPro 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Declarar Gasóleo Profesional"
   ClientHeight    =   3285
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7005
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3285
   ScaleWidth      =   7005
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdRecupera 
      Caption         =   "Recupera errores"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   2520
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Salir"
      Height          =   375
      Left            =   4800
      TabIndex        =   2
      Top             =   2520
      Width           =   1815
   End
   Begin VB.CommandButton cmdDeclarar 
      Caption         =   "Realizar la declaración"
      Height          =   375
      Left            =   2010
      TabIndex        =   1
      Top             =   2520
      Width           =   2415
   End
   Begin VB.Label lblinf 
      Caption         =   "Información de proceso"
      Height          =   285
      Left            =   120
      TabIndex        =   3
      Top             =   1710
      Width           =   6735
   End
   Begin VB.Label Label1 
      Caption         =   $"frmDecGasPro.frx":0000
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   1245
      Left            =   120
      TabIndex        =   0
      Top             =   210
      Width           =   6735
   End
End
Attribute VB_Name = "frmDecGasPro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Public Conn As ADODB.Connection
Private WithEvents gdgp As GestorDeclaracionesGasoleoProf
Attribute gdgp.VB_VarHelpID = -1

Private Sub cmdDeclarar_Click()
    Dim sql As String
    Set gdgp = New GestorDeclaracionesGasoleoProf
    If Not gdgp.quedaPorDeclarar(Conn, Now) Then
        MsgBox "No hay registros de Gasóleo Profesional para declarar", vbInformation
    Else
        If Not gdgp.declaraGasoleoProfesional(Conn, False) Then ' al loro el indicador ha de ser FALSE sin no las declaraciones se pasan en TEST
            MsgBox "Se ha producido un error durante la declaración del gasóleo, no se han generado los registros correspondientes", vbExclamation
            Exit Sub
        End If
    End If
    '-- En cualquier caso se intentan enviar las declaraciones por si acaso.
    If Not gdgp.enviaDeclaracionesPendientes(Conn) Then
        MsgBox "Se ha producido un error durante el envio de las declaraciones", vbExclamation
        Exit Sub
    Else
        MsgBox "Proceso realizado correctamente, todo lo pendiente se encuentra declarado.", vbInformation
        Unload Me
    End If
End Sub

Private Sub cmdRecupera_Click()
    Set gdgp = New GestorDeclaracionesGasoleoProf
    gdgp.recuperacionErrones Conn, "c:\aeat\RespuestasErroneas"
    MsgBox "Proceso finalizado consulte C:\AEAT\RespuestasRecuperadas"
    Unload Me
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    If Conn Is Nothing Then
        Set Conn = New Connection
        Conn.CursorLocation = adUseServer
        Conn.ConnectionString = "Provider=MSDASQL.1;Persist Security Info=False;Data Source=Arigasol" & _
                                 ";UID=root;PWD=aritel"
        Conn.Open
    End If
    lblInf.visible = False
End Sub

Private Sub gdGP_procesando(registro As Integer)
    lblInf.Caption = "Procesando registro " & CStr(registro)
    lblInf.Refresh
    DoEvents
End Sub


