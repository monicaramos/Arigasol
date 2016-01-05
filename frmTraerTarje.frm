VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmTraerTarje 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Busqueda Tarjetas de Clientes"
   ClientHeight    =   3375
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7350
   Icon            =   "frmTraerTarje.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3375
   ScaleWidth      =   7350
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   0
      Left            =   1080
      MaxLength       =   13
      TabIndex        =   0
      Tag             =   "año del Folleto|N|N|||follviaj|anyfovia|||"
      Top             =   480
      Width           =   1305
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   1350
      Left            =   120
      TabIndex        =   4
      Top             =   1320
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   2381
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   4920
      TabIndex        =   1
      Top             =   2865
      Width           =   1035
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   6180
      TabIndex        =   2
      Top             =   2865
      Width           =   1035
   End
   Begin VB.Label Label1 
      Caption         =   "Datos del Cliente"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   360
      TabIndex        =   6
      Top             =   1080
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Tarjeta"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   16
      Left            =   360
      TabIndex        =   5
      Top             =   480
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "Seleccione la tarjeta que desee buscar:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   240
      TabIndex        =   3
      Top             =   120
      Width           =   6015
   End
End
Attribute VB_Name = "frmTraerTarje"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Event Actualizar(vValor As Integer)


Public CodigoActual As String
Public Event DatoSeleccionado(CadenaSeleccion As String)
Public DatosADevolverBusqueda As String    'Tendra el nº de text que quiere que devuelva, empipados

Private Sub cmdAceptar_Click()
Dim cad As Integer, cadAux As String
Dim i As Integer
Dim NumF As Long
Dim J As Integer
Dim Aux As String

    On Error Resume Next
    
    Screen.MousePointer = vbHourglass
    
    If Me.ListView1.ListItems(1).Checked Then
        cad = Me.ListView1.ListItems(1)
    End If
    
    RaiseEvent Actualizar(cad)
    
    Screen.MousePointer = vbDefault
    Unload Me
    
    If Err.Number <> 0 Then Err.Clear
    
End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Text1(0).Text = ""
    PonerFoco Text1(0)
End Sub

Private Sub CargarListView()
Dim SQL As String
Dim Rs As ADODB.Recordset
Dim ItmX As ListItem

    On Error GoTo ECargar

    'Los encabezados
    ListView1.ColumnHeaders.Clear
    Me.ListView1.ListItems.Clear
    ListView1.ColumnHeaders.Add , , "Código", 1000
    ListView1.ColumnHeaders.Add , , "Nombre del Cliente", 3000
    ListView1.ColumnHeaders.Add , , "Nombre de la Tarjeta", 3000
    
    SQL = "select starje.codsocio, ssocio.nomsocio, starje.nomtarje from starje, ssocio"
    SQL = SQL & " WHERE starje.numtarje=" & Text1(0).Text & " AND starje.codsocio = ssocio.codsocio"
    SQL = SQL & " ORDER BY codsocio "
    
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    While Not Rs.EOF
        Set ItmX = ListView1.ListItems.Add
        ItmX.Text = Format(Rs.Fields(0).Value, "000000")
        
        ItmX.Checked = True
        
        ItmX.SubItems(1) = DBLet(Rs.Fields(1).Value, "T")
        ItmX.SubItems(2) = DBLet(Rs.Fields(2).Value, "T")
        Rs.MoveNext
    Wend
    Rs.Close
    Set Rs = Nothing
    
ECargar:
    If Err.Number <> 0 Then MuestraError Err.Number, "Cargar tarjetas.", Err.Description
End Sub

Private Sub Text1_GotFocus(Index As Integer)
    ConseguirFoco Text1(Index), 3
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub Text1_LostFocus(Index As Integer)
    If Text1(Index).Text <> "" Then
       Text1(0).Text = Format(Text1(0).Text, "0000000000000")
       CargarListView
    End If
End Sub

Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, 0, cerrar
    If cerrar Then Unload Me
End Sub

