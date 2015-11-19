VERSION 5.00
Begin VB.Form frmCRMReclamas 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Reclamaciones tesoreria(Arimoney)"
   ClientHeight    =   7305
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5985
   Icon            =   "frmCRMReclamas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7305
   ScaleWidth      =   5985
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox Combo1 
      Appearance      =   0  'Flat
      Height          =   315
      ItemData        =   "frmCRMReclamas.frx":000C
      Left            =   4560
      List            =   "frmCRMReclamas.frx":000E
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Tag             =   "Tipo de Cliente|N|N|||shcocob|carta|||"
      Top             =   3360
      Width           =   1245
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   375
      Index           =   1
      Left            =   4800
      TabIndex        =   10
      Top             =   6840
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Aceptar"
      Height          =   375
      Index           =   0
      Left            =   3720
      TabIndex        =   9
      Top             =   6840
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   2655
      Index           =   7
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   8
      Text            =   "frmCRMReclamas.frx":0010
      Top             =   3960
      Width           =   5535
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   6
      Left            =   2520
      MaxLength       =   10
      TabIndex        =   6
      Text            =   "Text1"
      Top             =   3360
      Width           =   1425
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   5
      Left            =   240
      MaxLength       =   10
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   3360
      Width           =   1425
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   4
      Left            =   3480
      MaxLength       =   3
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   2280
      Width           =   465
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   3
      Left            =   2040
      MaxLength       =   10
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   2280
      Width           =   1425
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   2
      Left            =   780
      MaxLength       =   10
      TabIndex        =   2
      Text            =   "1"
      Top             =   2280
      Width           =   1245
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   1
      Left            =   240
      MaxLength       =   1
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   2280
      Width           =   495
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H80000018&
      Height          =   315
      Index           =   0
      Left            =   1800
      Locked          =   -1  'True
      MaxLength       =   40
      TabIndex        =   13
      Top             =   1200
      Width           =   4005
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   0
      Left            =   240
      Locked          =   -1  'True
      MaxLength       =   10
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   1200
      Width           =   1425
   End
   Begin VB.Label Label1 
      Caption         =   "Envio"
      Height          =   255
      Index           =   8
      Left            =   4590
      TabIndex        =   24
      Top             =   3120
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   375
      Left            =   5520
      Top             =   120
      Width           =   375
   End
   Begin VB.Image imgFecha 
      Height          =   240
      Index           =   5
      Left            =   1320
      Picture         =   "frmCRMReclamas.frx":0016
      ToolTipText     =   "Buscar fecha"
      Top             =   3120
      Width           =   240
   End
   Begin VB.Image imgFecha 
      Height          =   240
      Index           =   3
      Left            =   3000
      Picture         =   "frmCRMReclamas.frx":00A1
      ToolTipText     =   "Buscar fecha"
      Top             =   2040
      Width           =   240
   End
   Begin VB.Label Label1 
      Caption         =   "Observaciones"
      Height          =   255
      Index           =   7
      Left            =   240
      TabIndex        =   23
      Top             =   3720
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Importe"
      Height          =   255
      Index           =   6
      Left            =   2550
      TabIndex        =   22
      Top             =   3120
      Width           =   1395
   End
   Begin VB.Label Label1 
      Caption         =   "F. reclamacion"
      Height          =   255
      Index           =   5
      Left            =   240
      TabIndex        =   21
      Top             =   3120
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Vto"
      Height          =   255
      Index           =   4
      Left            =   3600
      TabIndex        =   20
      Top             =   2040
      Width           =   375
   End
   Begin VB.Label Label1 
      Caption         =   "Fecha fact"
      Height          =   255
      Index           =   3
      Left            =   2040
      TabIndex        =   19
      Top             =   2040
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Factura"
      Height          =   255
      Index           =   2
      Left            =   780
      TabIndex        =   18
      Top             =   2040
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "Datos reclamación"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   300
      Index           =   1
      Left            =   120
      TabIndex        =   17
      Top             =   2760
      Width           =   5505
   End
   Begin VB.Label Label3 
      Caption         =   "Datos vencimiento"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   300
      Index           =   0
      Left            =   120
      TabIndex        =   16
      Top             =   1680
      Width           =   5505
   End
   Begin VB.Label Label2 
      Caption         =   "Datos cliente"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   300
      Left            =   120
      TabIndex        =   15
      Top             =   120
      Width           =   5505
   End
   Begin VB.Label Label1 
      Caption         =   "Serie"
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   14
      Top             =   2040
      Width           =   375
   End
   Begin VB.Label Label1 
      Caption         =   "Cuenta contable"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   12
      Top             =   960
      Width           =   1935
   End
   Begin VB.Label LabelCRM 
      Caption         =   "aqui ira el nomclient"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   420
      Left            =   240
      TabIndex        =   11
      Top             =   480
      Width           =   5505
   End
End
Attribute VB_Name = "frmCRMReclamas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Intercambio As String  ' codig|nomclien|codmacta|nommacta| estos ultimos si hiciera falta
Private codigo2  As Long
Private WithEvents frmC As frmCal
Attribute frmC.VB_VarHelpID = -1

Dim PrimeraVez As Boolean
Dim SQL As String
Dim i As Integer

Private Sub Command1_Click(Index As Integer)
    If Index = 0 Then
        If text1(5).Text = "" Or text1(6).Text = "" Or Combo1.ListIndex = -1 Then
            MsgBox "Campos obligatorios: fecha reclamacion, importe y envio", vbExclamation
            Exit Sub
        End If
        'Acciones...
        i = 0
        If codigo2 < 0 Then
            'NUEVO
            i = 1 'por si da error reestablacer el codigo2 a menos1
            'shcocob (codigo,numserie,codfaccl,fecfaccl,numorden,impvenci,codmacta,nommacta,carta,fecreclama,observaciones)
            SQL = DevuelveDesdeBDNew(cConta, "shcocob", "max(codigo)", "1", "1")
            If SQL = "" Then SQL = "0"
            codigo2 = Val(SQL) + 1
            SQL = "INSERT INTO shcocob (codigo,numserie,codfaccl,fecfaccl,numorden,impvenci,codmacta,nommacta,carta,"
            SQL = SQL & "fecreclama,observaciones) VALUES (" & codigo2 & ","
            SQL = SQL & DBSet(text1(1).Text, "T", "S") & "," '
            SQL = SQL & DBSet(text1(2).Text, "N", "S") & "," '= DBLet(miRsAux!Codfaccl, "T")
            SQL = SQL & DBSet(text1(3).Text, "F", "S") & "," '= DBLet(miRsAux!fecfaccl, "F")
            SQL = SQL & DBSet(text1(4).Text, "N", "S") & "," '= DBLet(miRsAux!numorden, "T")
            
            SQL = SQL & DBSet(text1(6).Text, "N") & "," ' z'= DBLet(miRsAux!ImpVenci, "N")
            SQL = SQL & DBSet(text1(0).Text, "T", "S") & "," ' = miRsAux!Codmacta
            SQL = SQL & DBSet(text2(0).Text, "T", "S") & "," & DBSet(Combo1.ListIndex, "N") & "," '0," ' = miRsAux!nommacta  Y  CARTA que le pondre un 0
            SQL = SQL & DBSet(text1(5).Text, "F") & "," ' = DBLet(miRsAux!fecreclama, "F")
            SQL = SQL & DBSet(text1(7).Text, "T", "S") & ")" ' = DBLet(miRsAux!Observaciones, "T")
            
        Else
            'MODIFICAR
            
            SQL = DBSet(text1(7).Text, "T")
            SQL = "UPDATE shcocob set observaciones = " & SQL & " WHERE codigo = " & codigo2
        End If
        
        ConnConta.Execute SQL
    End If
    Unload Me
End Sub

Private Sub Form_Activate()
    If Not PrimeraVez Then Exit Sub
    PrimeraVez = False
    
    If codigo2 > 0 Then
        Set miRsAux = New ADODB.Recordset
        SQL = "Select * from shcocob where codigo = " & codigo2
        miRsAux.Open SQL, ConnConta, adOpenKeyset, adLockPessimistic, adCmdText
        If Not miRsAux.EOF Then
            text1(0).Text = miRsAux!codmacta
            text2(0).Text = miRsAux!nommacta
            text1(1).Text = DBLet(miRsAux!numserie, "T")
            text1(2).Text = DBLet(miRsAux!codfaccl, "T")
            text1(3).Text = DBLet(miRsAux!fecfaccl, "F")
            text1(4).Text = DBLet(miRsAux!numorden, "T")
            text1(5).Text = DBLet(miRsAux!fecreclama, "F")
            text1(6).Text = DBLet(miRsAux!ImpVenci, "N")
            text1(7).Text = DBLet(miRsAux!Observaciones, "T")
            
            PosicionarCombo Combo1, miRsAux!carta
            
        Else
            MsgBox "Imposible abrir reclamacion cod: " & codigo2, vbExclamation
            codigo2 = -1
            limpiar Me
            Me.Command1(0).visible = False
        End If
        miRsAux.Close
    End If
    
    If codigo2 < 0 Then
        For i = 1 To 7
            text1(i).Text = ""
        Next i
        text1(5).Text = Format(Now, "dd/mm/yyyy")
        Combo1.ListIndex = 2
    End If
    text1(0).Locked = True
    For i = 1 To 6
        text1(i).Locked = codigo2 >= 0
    Next
    
    Combo1.Locked = (codigo2 >= 0)
    
    Me.imgFecha(3).visible = codigo2 = -1
    Me.imgFecha(5).visible = codigo2 = -1
    If codigo2 >= 0 Then
        PonerFoco text1(7)
    Else
        PonerFoco text1(1)
    End If
End Sub

Private Sub Form_Load()

'    Me.Icon = frmPpal.Icon
    PrimeraVez = True
    codigo2 = Val(RecuperaValor(Intercambio, 1))
    Me.LabelCRM.Caption = RecuperaValor(Intercambio, 2)
    If codigo2 < 0 Then
        text1(0).Text = RecuperaValor(Intercambio, 3)
        text2(0).Text = RecuperaValor(Intercambio, 4)
    End If
    Image1.Picture = frmPpal.imgListComun.ListImages(4).Picture '46
    
    CargaCombo
    
End Sub

Private Sub frmC_Selec(vFecha As Date)
    ' *** repasar si el camp es txtAux o Text1 ***
    text1(CByte(imgFecha(3).Tag)).Text = Format(vFecha, "dd/mm/yyyy") '<===
    ' ********************************************
End Sub

Private Sub Image1_Click()
    SQL = "Reclamaciones" & vbCrLf & String(40, "-") & vbCrLf & vbCrLf
    SQL = SQL & "Si esta modificando solo se le permite cambiar las observaciones." & vbCrLf
    SQL = SQL & "Si es nueva son obligatorios los campos: fecha reclamacion, importe y envio" & vbCrLf
    MsgBox SQL, vbInformation
End Sub

Private Sub imgFecha_Click(Index As Integer)
    Dim esq As Long
    Dim dalt As Long
    Dim menu As Long
    Dim obj As Object


    Set frmC = New frmCal
    
    esq = imgFecha(Index).Left
    dalt = imgFecha(Index).Top
        
    Set obj = imgFecha(Index).Container
      
    While imgFecha(Index).Parent.Name <> obj.Name
          esq = esq + obj.Left
          dalt = dalt + obj.Top
          Set obj = obj.Container
    Wend
    
    menu = Me.Height - Me.ScaleHeight 'ací tinc el heigth del menú i de la toolbar

    frmC.Left = esq + imgFecha(Index).Parent.Left + 30
    frmC.Top = dalt + imgFecha(Index).Parent.Top + imgFecha(Index).Height + menu - 40

    imgFecha(3).Tag = Index

    SQL = ""
'    Set frmC = New frmCal
    frmC.NovaData = Now
    If text1(Index).Text <> "" Then frmC.NovaData = CDate(text1(Index).Text)
    frmC.Show vbModal
    If SQL <> "" Then text1(Index).Text = SQL
    SQL = ""
End Sub

Private Sub Text1_GotFocus(Index As Integer)
    
    If Not text1(Index).Locked And Index <> 7 Then ConseguirFoco text1(Index), 3
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpressGnral KeyAscii, 3, False
End Sub

Private Sub Text1_LostFocus(Index As Integer)
    text1(Index).Text = Trim(text1(Index).Text)
    If text1(Index).Text = "" Then Exit Sub
    
    Select Case Index
    Case 2, 4
        If Not PonerFormatoEntero(text1(Index)) Then text1(Index).Text = ""
    
    Case 3, 5
        PonerFormatoFecha text1(Index)
    
    Case 6
        If Not PonerFormatoDecimal(text1(Index), 1) Then text1(Index).Text = ""
    
    End Select
    
    
End Sub


Private Sub CargaCombo()
    Combo1.Clear
    'Conceptos
'    Set miRsAux = New ADODB.Recordset
'    miRsAux.Open "Select * from stipoformapago order by descformapago", Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
'    While Not miRsAux.EOF
'        Combo1.AddItem miRsAux!descformapago
'        Combo1.ItemData(Combo1.NewIndex) = miRsAux!tipoformapago
'        miRsAux.MoveNext
'    Wend
'    miRsAux.Close
'    Set miRsAux = Nothing
        Combo1.AddItem "Carta"
        Combo1.ItemData(Combo1.NewIndex) = 0
        Combo1.AddItem "Email"
        Combo1.ItemData(Combo1.NewIndex) = 1
        Combo1.AddItem "Teléfono"
        Combo1.ItemData(Combo1.NewIndex) = 2
End Sub

