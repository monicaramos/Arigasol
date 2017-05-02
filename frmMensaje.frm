VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMensaje 
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   ClientHeight    =   5700
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9015
   LinkTopic       =   "Form1"
   ScaleHeight     =   5700
   ScaleWidth      =   9015
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameTarjetasLibres 
      BorderStyle     =   0  'None
      Height          =   5160
      Left            =   0
      TabIndex        =   19
      Top             =   0
      Width           =   5280
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   1
         Left            =   3330
         MaxLength       =   13
         TabIndex        =   24
         Tag             =   "año del Folleto|N|N|||follviaj|anyfovia|||"
         Top             =   990
         Width           =   1350
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   0
         Left            =   1170
         MaxLength       =   13
         TabIndex        =   23
         Tag             =   "año del Folleto|N|N|||follviaj|anyfovia|||"
         Top             =   990
         Width           =   1350
      End
      Begin VB.CommandButton CmdAcep 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   2385
         TabIndex        =   26
         Top             =   4680
         Width           =   1035
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   3645
         TabIndex        =   28
         Top             =   4680
         Width           =   1035
      End
      Begin MSComctlLib.ListView ListView3 
         Height          =   3165
         Left            =   210
         TabIndex        =   20
         Top             =   1395
         Width           =   4545
         _ExtentX        =   8017
         _ExtentY        =   5583
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         HideColumnHeaders=   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H00800000&
         BorderWidth     =   3
         Height          =   5055
         Left            =   30
         Top             =   90
         Width           =   5190
      End
      Begin VB.Label Label1 
         Caption         =   "Búsqueda de Tarjetas Libres"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   5
         Left            =   225
         TabIndex        =   27
         Top             =   270
         Width           =   4275
      End
      Begin VB.Label Label1 
         Caption         =   "Tarjeta"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   4
         Left            =   270
         TabIndex        =   25
         Top             =   765
         Width           =   945
      End
      Begin VB.Label Label1 
         Caption         =   "Hasta"
         Height          =   255
         Index           =   3
         Left            =   2790
         TabIndex        =   22
         Top             =   1035
         Width           =   675
      End
      Begin VB.Label Label1 
         Caption         =   "Desde"
         Height          =   255
         Index           =   0
         Left            =   585
         TabIndex        =   21
         Top             =   1035
         Width           =   675
      End
   End
   Begin VB.Frame FrameClientesLibres 
      BorderStyle     =   0  'None
      Height          =   2865
      Left            =   0
      TabIndex        =   30
      Top             =   0
      Width           =   5280
      Begin VB.CommandButton CmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   3600
         TabIndex        =   34
         Top             =   2025
         Width           =   1035
      End
      Begin VB.CommandButton CmdAcept 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   2340
         TabIndex        =   33
         Top             =   2025
         Width           =   1035
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   3
         Left            =   1890
         MaxLength       =   10
         TabIndex        =   32
         Tag             =   "Código de cliente|N|N|0|999999|ssocio|codsocio|000000|S|"
         Top             =   720
         Width           =   1350
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0080FFFF&
         Enabled         =   0   'False
         Height          =   285
         Index           =   2
         Left            =   1890
         MaxLength       =   10
         TabIndex        =   31
         Tag             =   "año del Folleto|N|N|||follviaj|anyfovia|||"
         Top             =   1215
         Width           =   1350
      End
      Begin VB.Label Label1 
         Caption         =   "Código siguiente:"
         Height          =   255
         Index           =   8
         Left            =   405
         TabIndex        =   37
         Top             =   1260
         Width           =   1485
      End
      Begin VB.Label Label1 
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
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   7
         Left            =   405
         TabIndex        =   36
         Top             =   720
         Width           =   945
      End
      Begin VB.Label Label1 
         Caption         =   "Búsqueda de Clientes Libres"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   6
         Left            =   225
         TabIndex        =   35
         Top             =   270
         Width           =   4275
      End
      Begin VB.Shape Shape4 
         BorderColor     =   &H00800000&
         BorderWidth     =   3
         Height          =   2760
         Left            =   30
         Top             =   90
         Width           =   5190
      End
   End
   Begin VB.Frame FrameCobrosPtes 
      Height          =   3090
      Left            =   135
      TabIndex        =   0
      Top             =   945
      Width           =   8295
      Begin VB.CommandButton CmdCancelarCobros 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   6885
         TabIndex        =   3
         Top             =   2385
         Visible         =   0   'False
         Width           =   1035
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   1815
         Left            =   225
         TabIndex        =   2
         Top             =   450
         Width           =   7785
         _ExtentX        =   13732
         _ExtentY        =   3201
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.CommandButton CmdAceptarCobros 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   5760
         TabIndex        =   1
         Top             =   2400
         Width           =   1035
      End
      Begin VB.Label Label3 
         Caption         =   "Cobros con Errores"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   270
         TabIndex        =   29
         Top             =   180
         Width           =   3585
      End
      Begin VB.Label Label1 
         Caption         =   "Label2"
         Height          =   345
         Index           =   1
         Left            =   450
         TabIndex        =   4
         Top             =   1470
         Width           =   3555
      End
   End
   Begin VB.Frame FrameErrores 
      Height          =   5505
      Left            =   60
      TabIndex        =   5
      Top             =   90
      Width           =   8835
      Begin VB.CommandButton CmdSalir 
         Caption         =   "&Salir"
         Height          =   375
         Left            =   6930
         TabIndex        =   6
         Top             =   4830
         Width           =   1035
      End
      Begin MSComctlLib.ListView ListView2 
         Height          =   4155
         Left            =   210
         TabIndex        =   7
         Top             =   540
         Width           =   7785
         _ExtentX        =   13732
         _ExtentY        =   7329
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.Label Label2 
         Caption         =   "Errores de Comprobación"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   270
         TabIndex        =   9
         Top             =   210
         Width           =   3585
      End
      Begin VB.Label Label1 
         Caption         =   "Label2"
         Height          =   345
         Index           =   2
         Left            =   450
         TabIndex        =   8
         Top             =   1470
         Width           =   3555
      End
   End
   Begin VB.Frame frameAcercaDE 
      BorderStyle     =   0  'None
      ForeColor       =   &H00C00000&
      Height          =   4545
      Left            =   90
      TabIndex        =   10
      Top             =   240
      Width           =   5385
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   3840
         TabIndex        =   11
         Top             =   3900
         Width           =   1035
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fax: 96 3420938"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   195
         Left            =   3180
         TabIndex        =   18
         Top             =   3540
         Width           =   1440
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tfno: 902 888 878"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   195
         Left            =   780
         TabIndex        =   17
         Top             =   3540
         Width           =   1590
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Arigasol"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   36
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   915
         Left            =   240
         TabIndex        =   16
         Top             =   360
         Width           =   4335
      End
      Begin VB.Label Label25 
         BackStyle       =   0  'Transparent
         Caption         =   "4"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   81.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   1725
         Left            =   3780
         TabIndex        =   15
         Top             =   60
         Width           =   1350
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "46007 - VALENCIA"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   195
         Left            =   3240
         TabIndex        =   14
         Top             =   3120
         Width           =   1620
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "C/ Uruguay, 11 - Despacho 101"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   195
         Left            =   300
         TabIndex        =   13
         Top             =   3120
         Width           =   2730
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Label13"
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
         Left            =   0
         TabIndex        =   12
         Top             =   1200
         Width           =   3795
      End
      Begin VB.Image Image2 
         Height          =   540
         Left            =   1740
         Top             =   2460
         Width           =   2880
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00B48246&
         BorderWidth     =   5
         FillColor       =   &H00E0E0E0&
         FillStyle       =   0  'Solid
         Height          =   4425
         Left            =   90
         Top             =   60
         Width           =   5250
      End
   End
End
Attribute VB_Name = "frmMensaje"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public OpcionMensaje As Byte

'variables que se pasan con valor al llamar al formulario de zoom desde otro formulario


Public pTitulo As String



Private Sub CmdAcep_Click()
    CargarTarjetasLibres
End Sub

Private Sub CmdAcept_Click()
Dim SQL As String
Dim Rs As ADODB.Recordset
Dim Anterior As Currency
Dim Siguiente As Currency
Dim Encontrado As Boolean

    If Text1(3).Text = "" Then Text1(3).Text = 0
    
    Set Rs = New ADODB.Recordset
    SQL = "select codsocio from ssocio where codsocio >= " & DBSet(Text1(3).Text, "N")
    Rs.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText

    Anterior = CCur(Text1(3).Text)
    Encontrado = False
    While Not Rs.EOF And Not Encontrado
        If Anterior = DBLet(Rs.Fields(0).Value, "N") Then
            Anterior = Anterior + 1
        Else
            Siguiente = Anterior
            Encontrado = True
        End If
        Rs.MoveNext
    Wend
    
    Set Rs = Nothing
    Text1(2).Text = Format(Siguiente, "000000")
    
    cmdCancel.SetFocus
End Sub

Private Sub cmdAceptar_Click()
    Unload Me
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdCancelar_Click()
    'salimos y no hacemos nada
    Unload Me
End Sub

Private Sub CmdAceptarCobros_Click()
     Unload Me
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Command2_Click()

End Sub

Private Sub Form_Activate()
    PonerFocoBtn Me.cmdAceptar
    If OpcionMensaje = 4 Then PonerFoco Text1(0)
End Sub


Private Sub Form_Load()
    Me.Shape1.Width = Me.Width - 30
    Me.Shape1.Height = Me.Height - 30

    'obtener el campo correspondiente y mostrarlo en el text
    
    Label1(1).Caption = pTitulo

    If OpcionMensaje <= 3 Then ' Errores al hacer comprobaciones
        PonerFrameCobrosPtesVisible True, 1000, 2000
        CargarListaErrComprobacion
        Me.Caption = "Errores de Comprobacion: "
        PonerFocoBtn Me.CmdSalir
    End If
    

    If OpcionMensaje = 10 Then  'Errores al contabilizar facturas
        PonerFrameCobrosPtesVisible True, 1000, 2000
        CargarListaErrContab
        Me.Caption = "Facturas NO contabilizadas: "
        PonerFocoBtn Me.CmdAceptarCobros
    End If

    If OpcionMensaje = 6 Then
        PonerFrameCobrosPtesVisible True, 1000, 2000
        CargaImagen
'        Me.Caption = "Acerca de ....."
'        w = Me.frameAcercaDE.Width
'        h = Me.frameAcercaDE.Height
        Me.frameAcercaDE.visible = True
        Label13.Caption = "Versión:  " & App.Major & "." & App.Minor & "." & App.Revision & " "
    End If
    
    If OpcionMensaje = 4 Then ' tarjetas libres
        PonerFrameCobrosPtesVisible True, 1000, 4000
    End If
    
    If OpcionMensaje = 5 Then ' clientes libres
        PonerFrameCobrosPtesVisible True, 1000, 2000
    End If

End Sub

Private Sub CargarListaErrContab()
'Muestra la lista Detallada de Facturas que dieron error al contabilizar
'en un ListView
Dim Rs As ADODB.Recordset
Dim ItmX As ListItem
Dim SQL As String

    On Error GoTo ECargarList

    SQL = " SELECT  * "
    SQL = SQL & " FROM tmperrfac "
    
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not Rs.EOF Then
        ListView1.Height = 4500
        ListView1.Width = 7400
        ListView1.Left = 500
        ListView1.Top = 500

        'Los encabezados
        ListView1.ColumnHeaders.Clear

        If Rs.Fields(0).Name = "codprove" Then
            'Facturas de Compra
             ListView1.ColumnHeaders.Add , , "Prove.", 700
        Else 'Facturas de Venta
            ListView1.ColumnHeaders.Add , , "Tipo", 600
        End If
        ListView1.ColumnHeaders.Add , , "Factura", 1000, 1
        ListView1.ColumnHeaders.Add , , "Fecha", 1100, 1
        ListView1.ColumnHeaders.Add , , "Error", 4620
    
        While Not Rs.EOF
            Set ItmX = ListView1.ListItems.Add
            'El primer campo será codtipom si llamamos desde Ventas
            ' y será codprove si llamamos desde Compras
            ItmX.Text = Rs.Fields(0).Value
            ItmX.SubItems(1) = Format(Rs!numfactu, "0000000")
            ItmX.SubItems(2) = Rs!Fecfactu
            ItmX.SubItems(3) = Rs!Error
            Rs.MoveNext
        Wend
    End If
    Rs.Close
    Set Rs = Nothing

ECargarList:
    If Err.Number <> 0 Then
        Err.Clear
    End If
End Sub

Private Sub PonerFrameCobrosPtesVisible(visible As Boolean, ByRef h As Integer, ByRef w As Integer)
'Pone el Frame de Cobros Pendientes Visible y Ajustado al Formulario, y visualiza los controles
'necesario para el Informe

    h = 4600
        
    Select Case OpcionMensaje
        Case 0, 1, 2, 3
            h = 6000
            w = 9200
'            Me.Label1(0).Top = 4800
'            Me.Label1(0).Left = 3400
            Me.CmdSalir.Caption = "&Salir"
            PonerFrameVisible Me.FrameErrores, visible, h, w
            Me.frameAcercaDE.visible = False
            Me.FrameCobrosPtes.visible = False
            Me.FrameTarjetasLibres.visible = False
            Me.FrameClientesLibres.visible = False
            
        
        Case 10  'Errores al contabilizar facturas
            h = 6000
            w = 8400
            Me.CmdAceptarCobros.Top = 5300
            Me.CmdAceptarCobros.Left = 4900
            '++monica
            PonerFrameVisible Me.FrameCobrosPtes, visible, h, w
            Me.frameAcercaDE.visible = False
            Me.FrameCobrosPtes.visible = True
            Me.FrameTarjetasLibres.visible = False
            Me.FrameErrores.visible = False
            Me.FrameClientesLibres.visible = False
            
        Case 6 ' Acerca de
            h = 4485
            w = 5415
            Me.Width = w
            Me.Height = h
            Me.Shape1.Width = w
            Me.Shape1.Height = h
            Me.Shape1.Top = 0
            Me.Shape1.Left = 0
            Me.frameAcercaDE.visible = True
            Me.frameAcercaDE.Left = 5
            Me.frameAcercaDE.Top = 5
            Me.frameAcercaDE.Width = w - 5
            Me.frameAcercaDE.Height = h - 5
            Me.FrameCobrosPtes.visible = False
            Me.FrameErrores.visible = False
            Me.FrameTarjetasLibres.visible = False
            Me.FrameClientesLibres.visible = False
        
'            PonerFrameVisible Me.frameAcercaDE, visible, h - 20, w - 20

            Exit Sub
            
        Case 4
            Me.FrameTarjetasLibres.visible = True
            h = 5160
            w = 5280
            Me.FrameTarjetasLibres.Width = w
            Me.FrameTarjetasLibres.Height = h
            Me.Width = w
            Me.Height = h
            Me.FrameTarjetasLibres.Top = 0
            Me.FrameTarjetasLibres.Left = 0
            Me.frameAcercaDE.visible = False
            Me.FrameCobrosPtes.visible = False
            Me.FrameErrores.visible = False
            Me.FrameClientesLibres.visible = False
            Me.Shape2.Width = w
            Me.Shape2.Height = h
            Me.Shape2.Top = 0
            Me.Shape2.Left = 0
            'Los encabezados
            ListView3.ColumnHeaders.Clear
            ListView3.ColumnHeaders.Add , , "Tarjeta", 4500
            Text1(0).Text = ""
            Text1(1).Text = ""
'            CmdAcep_Click
            PonerFoco Text1(0)

        Case 5 ' clientes libres
            Me.FrameClientesLibres.visible = True
            h = 2730
            w = 5160
            Me.FrameClientesLibres.Width = w
            Me.FrameClientesLibres.Height = h
            Me.Width = w
            Me.Height = h
            Me.FrameClientesLibres.Top = 0
            Me.FrameClientesLibres.Left = 0
            Me.frameAcercaDE.visible = False
            Me.FrameCobrosPtes.visible = False
            Me.FrameErrores.visible = False
            Me.FrameTarjetasLibres.visible = False
            Me.Shape4.Width = w
            Me.Shape4.Height = h
            Me.Shape4.Top = 0
            Me.Shape4.Left = 0
            'Los encabezados
            PonerFoco Text1(3)
       


    End Select
            
    
End Sub


Private Sub CargarListaErrComprobacion()
'Muestra la lista Detallada de Facturas que dieron error al contabilizar
'en un ListView
Dim Rs As ADODB.Recordset
Dim ItmX As ListItem
Dim SQL As String

    On Error GoTo ECargarListErrComprobacion

    SQL = " SELECT  * "
    SQL = SQL & " FROM tmperrcomprob "
    
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not Rs.EOF Then
'        ListView1.Height = 4500
'        ListView1.Width = 7400
'        ListView1.Left = 500
'        ListView1.Top = 500

        'Los encabezados
        ListView2.ColumnHeaders.Clear

        Select Case OpcionMensaje
            Case 1
                ListView2.ColumnHeaders.Add , , "Error en letra de serie", 6000
            Case 2
                ListView2.ColumnHeaders.Add , , "Error en cuentas contables", 6000
            Case 3
                ListView2.ColumnHeaders.Add , , "Error en tipos de iva", 6000
        
        End Select


'        ListView2.ColumnHeaders.Add , , "Error de comprobación", 5000
'
'        If RS.Fields(0).Name = "codprove" Then
'            'Facturas de Compra
'             ListView1.ColumnHeaders.Add , , "Prove.", 700
'        Else 'Facturas de Venta
'            ListView1.ColumnHeaders.Add , , "Tipo", 600
'        End If
'        ListView1.ColumnHeaders.Add , , "Factura", 1000, 1
'        ListView1.ColumnHeaders.Add , , "Fecha", 1100, 1
'        ListView1.ColumnHeaders.Add , , "Error", 4620
    
        While Not Rs.EOF
            Set ItmX = ListView2.ListItems.Add
            'El primer campo será codtipom si llamamos desde Ventas
            ' y será codprove si llamamos desde Compras
            ItmX.Text = Rs.Fields(0).Value
'            ItmX.SubItems(1) = Format(RS!NumFactu, "0000000")
'            ItmX.SubItems(2) = RS!FecFactu
'            ItmX.SubItems(3) = RS!Error
            Rs.MoveNext
        Wend
    End If
    Rs.Close
    Set Rs = Nothing

ECargarListErrComprobacion:
    If Err.Number <> 0 Then
        Err.Clear
    End If
End Sub

Private Sub CargaImagen()
On Error Resume Next
     Image2.Picture = LoadPicture(App.path & "\logo.jpg")
'    Image2.Picture = LoadPicture(App.path & "\images\minilogo.bmp")
'    Image1.Picture = LoadPicture(App.path & "\fondon.gif")
    Err.Clear
End Sub


Private Sub CargarTarjetasLibres()
'Muestra la lista Detallada de Tarjetas Libres
Dim Rs As ADODB.Recordset
Dim ItmX As ListItem
Dim SQL As String
Dim J As Double
Dim I As Double

    On Error GoTo ECargarList

    ListView3.ListItems.Clear
    ListView3.ColumnHeaders.Clear
    ListView3.ColumnHeaders.Add , , "Tarjeta", 2000

    SQL = " SELECT  numtarje "
    SQL = SQL & " FROM starje where 1=1 "
    If Text1(0).Text <> "" Then SQL = SQL & " and numtarje >= " & DBSet(Text1(0).Text, "N")
    If Text1(1).Text <> "" Then SQL = SQL & " and numtarje <= " & DBSet(Text1(1).Text, "N")
    SQL = SQL & " order by numtarje"
    
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not Rs.EOF Then
        If Text1(0).Text = "" Then
            J = 1
        Else
            J = CCur(Text1(0).Text)
        End If
    
        While Not Rs.EOF
            For I = J To DBLet(Rs.Fields(0).Value, "N") - 1
                Set ItmX = ListView3.ListItems.Add
                ItmX.Text = Format(I, "0000000000000")
            Next I
            J = DBLet(Rs.Fields(0).Value, "N") + 1
            Rs.MoveNext
        Wend
        If Text1(1).Text <> "" Then
            For I = J To CCur(Text1(1).Text)
                Set ItmX = ListView3.ListItems.Add
                ItmX.Text = Format(I, "0000000000000")
            Next I
        Else
            For I = J To 9999999999999#
                Set ItmX = ListView3.ListItems.Add
                ItmX.Text = Format(I, "0000000000000")
            Next I
        End If
    Else
        If Text1(0).Text = "" Then
            J = 1
        Else
            J = CCur(Text1(0).Text)
        End If
        If Text1(1).Text <> "" Then
            For I = J To CCur(Text1(1).Text)
                Set ItmX = ListView3.ListItems.Add
                ItmX.Text = Format(I, "0000000000000")
            Next I
        Else
            For I = J To 9999999999999#
                Set ItmX = ListView3.ListItems.Add
                ItmX.Text = Format(I, "0000000000000")
            Next I
        End If
   
    End If
    Rs.Close
    Set Rs = Nothing

ECargarList:
    If Err.Number <> 0 Then
        Err.Clear
    End If
End Sub

Private Sub ListView3_DblClick()
    Me.pTitulo = ListView3.SelectedItem
    cmdCancelar_Click
End Sub

Private Sub Text1_GotFocus(Index As Integer)
    ConseguirFoco Text1(Index), 3
End Sub

Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    KEYdown KeyCode
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, 0, cerrar
    If cerrar Then Unload Me
End Sub

Private Sub Text1_LostFocus(Index As Integer)
Dim Cad As String, cadTipo As String 'tipo cliente
Dim dev As String
    'Quitar espacios en blanco por los lados
    Text1(Index).Text = Trim(Text1(Index).Text)
'    If txtCodigo(Index).Text = "" Then Exit Sub
    
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub

    Select Case Index
        Case 0, 1 'Tarjetas
            If Text1(Index).Text <> "" Then Text1(Index).Text = Format(Text1(Index).Text, "0000000000000")
            
            If Text1(0).Text <> "" And Text1(1).Text <> "" Then
                dev = CadenaDesdeHasta(Text1(0).Text, Text1(1).Text, "numtarje", "N")
                If dev = "Error" Then
                    PonerFoco Text1(0)
                End If
            End If
        
        Case 3 ' Codigo de Cliente
            PonerFormatoEntero Text1(3)

    End Select
End Sub

