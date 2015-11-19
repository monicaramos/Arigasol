VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmEstCliimp 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ventas por Cliente"
   ClientHeight    =   6315
   ClientLeft      =   45
   ClientTop       =   2430
   ClientWidth     =   7185
   Icon            =   "frmEstCliimp.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6315
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
      Height          =   6015
      Left            =   120
      TabIndex        =   10
      Top             =   120
      Width           =   6915
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         ForeColor       =   &H00972E0B&
         Height          =   945
         Left            =   3390
         TabIndex        =   26
         Top             =   4050
         Width           =   2835
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   6
            Left            =   600
            MaxLength       =   15
            TabIndex        =   7
            Top             =   420
            Width           =   1875
         End
         Begin VB.Label Label4 
            Caption         =   "Importe Factura superior a: "
            ForeColor       =   &H00972E0B&
            Height          =   195
            Index           =   3
            Left            =   120
            TabIndex        =   27
            Top             =   60
            Width           =   1995
         End
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   5
         Left            =   2745
         Locked          =   -1  'True
         TabIndex        =   21
         Text            =   "Text5"
         Top             =   3375
         Width           =   3135
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   4
         Left            =   2745
         Locked          =   -1  'True
         TabIndex        =   20
         Text            =   "Text5"
         Top             =   3000
         Width           =   3135
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   5
         Left            =   1845
         MaxLength       =   6
         TabIndex        =   5
         Top             =   3360
         Width           =   830
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   4
         Left            =   1845
         MaxLength       =   6
         TabIndex        =   4
         Top             =   3000
         Width           =   830
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   3
         Left            =   1845
         MaxLength       =   10
         TabIndex        =   3
         Tag             =   "Código Postal|T|S|||clientes|codposta|||"
         Top             =   2280
         Width           =   1050
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   2
         Left            =   1845
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
         TabIndex        =   9
         Top             =   5400
         Width           =   975
      End
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   3720
         TabIndex        =   8
         Top             =   5400
         Width           =   975
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   0
         Left            =   1845
         MaxLength       =   6
         TabIndex        =   0
         Top             =   840
         Width           =   830
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   1
         Left            =   1845
         MaxLength       =   6
         TabIndex        =   1
         Top             =   1215
         Width           =   830
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   2745
         Locked          =   -1  'True
         TabIndex        =   13
         Text            =   "Text5"
         Top             =   840
         Width           =   3135
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   2745
         Locked          =   -1  'True
         TabIndex        =   12
         Text            =   "Text5"
         Top             =   1215
         Width           =   3135
      End
      Begin VB.Frame Frame4 
         Caption         =   "Tipo Informe"
         ForeColor       =   &H00972E0B&
         Height          =   1000
         Left            =   600
         TabIndex        =   11
         Top             =   4080
         Width           =   2235
         Begin VB.OptionButton Option1 
            Caption         =   "Resumen"
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   25
            Top             =   600
            Width           =   1815
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Detalle"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   6
            Top             =   240
            Value           =   -1  'True
            Width           =   1815
         End
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   5
         Left            =   1545
         MouseIcon       =   "frmEstCliimp.frx":000C
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar colectivo"
         Top             =   3375
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   4
         Left            =   1560
         MouseIcon       =   "frmEstCliimp.frx":015E
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar colectivo"
         Top             =   3000
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
         Index           =   2
         Left            =   600
         TabIndex        =   24
         Top             =   2760
         Width           =   660
      End
      Begin VB.Label Label4 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   1
         Left            =   960
         TabIndex        =   23
         Top             =   3375
         Width           =   420
      End
      Begin VB.Label Label4 
         Caption         =   "Desde"
         Height          =   195
         Index           =   0
         Left            =   960
         TabIndex        =   22
         Top             =   3000
         Width           =   465
      End
      Begin VB.Label Label4 
         Caption         =   "Fecha factura"
         ForeColor       =   &H00972E0B&
         Height          =   255
         Index           =   16
         Left            =   600
         TabIndex        =   19
         Top             =   1680
         Width           =   1815
      End
      Begin VB.Label Label4 
         Caption         =   "Desde"
         Height          =   195
         Index           =   15
         Left            =   960
         TabIndex        =   18
         Top             =   1920
         Width           =   465
      End
      Begin VB.Label Label4 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   14
         Left            =   960
         TabIndex        =   17
         Top             =   2280
         Width           =   420
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   2
         Left            =   1530
         Picture         =   "frmEstCliimp.frx":02B0
         ToolTipText     =   "Buscar fecha"
         Top             =   1920
         Width           =   240
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   3
         Left            =   1530
         Picture         =   "frmEstCliimp.frx":033B
         ToolTipText     =   "Buscar fecha"
         Top             =   2280
         Width           =   240
      End
      Begin VB.Label Label4 
         Caption         =   "Desde"
         Height          =   195
         Index           =   13
         Left            =   960
         TabIndex        =   16
         Top             =   840
         Width           =   465
      End
      Begin VB.Label Label4 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   12
         Left            =   960
         TabIndex        =   15
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
         Left            =   600
         TabIndex        =   14
         Top             =   600
         Width           =   495
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   0
         Left            =   1560
         MouseIcon       =   "frmEstCliimp.frx":03C6
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar cliente"
         Top             =   840
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   1
         Left            =   1545
         MouseIcon       =   "frmEstCliimp.frx":0518
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar cliente"
         Top             =   1215
         Width           =   240
      End
   End
End
Attribute VB_Name = "frmEstCliimp"
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

Private WithEvents frmCol As frmManCoope 'Colectivo
Attribute frmCol.VB_VarHelpID = -1
Private WithEvents frmcli As frmManClien 'Clientes
Attribute frmcli.VB_VarHelpID = -1
Private WithEvents frmC As frmCal 'calendario fechas
Attribute frmC.VB_VarHelpID = -1

'GENERALES PARA PASARLE A CRYSTAL REPORT
Private cadFormula As String 'Cadena con la FormulaSelection para Crystal Report
Private cadParam As String 'Cadena con los parametros para Crystal Report
Private numParam As Byte 'Numero de parametros que se pasan a Crystal Report
Private cadSelect As String 'Cadena para comprobar si hay datos antes de abrir Informe
Private cadTitulo As String 'Titulo para la ventana frmImprimir
Private cadNombreRPT As String 'Nombre del informe

Dim indCodigo As Integer 'indice para txtCodigo
Dim indFrame As Single 'nº de frame en el que estamos
 
'Se inicializan para cada Informe (tabla de BD a la que hace referencia
Dim Tabla As String
Dim Codigo As String 'Código para FormulaSelection de Crystal Report
Dim TipCod As String
Dim Orden1 As String 'Campo de Ordenacion (por codigo) para Cristal Report
Dim Orden2 As String 'Campo de Ordenacion (por nombre) para Cristal Report

Dim PrimeraVez As Boolean


Dim Periodo1 As String
Dim Periodo2 As String
Dim Periodo3 As String
Dim Periodo4 As String


Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, 0, cerrar
    If cerrar Then Unload Me
End Sub

Private Sub cmdAceptar_Click()
Dim cDesde As String, cHasta As String 'cadena codigo Desde/Hasta
Dim nDesde As String, nHasta As String 'cadena Descripcion Desde/Hasta
Dim cadTABLA As String, cOrden As String
Dim I As Byte
InicializarVbles
    
    '========= PARAMETROS  =============================
    'Añadir el parametro de Empresa
    cadParam = cadParam & "|pEmpresa=""" & vEmpresa.nomEmpre & """|"
    numParam = numParam + 1
    
     '======== FORMULA  ====================================
    'Seleccionar registros de la empresa conectada
'    Codigo = "{" & tabla & ".codempre}=" & vEmpresa.codEmpre
'    If Not AnyadirAFormula(cadFormula, Codigo) Then Exit Sub
'    If Not AnyadirAFormula(cadSelect, Codigo) Then Exit Sub
    
    
    'D/H Cliente
    cDesde = Trim(txtCodigo(0).Text)
    cHasta = Trim(txtCodigo(1).Text)
    nDesde = txtNombre(0).Text
    nHasta = txtNombre(1).Text
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{" & Tabla & ".codsocio}"
        TipCod = "N"
        If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHcliente= """) Then Exit Sub
    End If
    
    'D/H Fecha factura
    cDesde = Trim(txtCodigo(2).Text)
    cHasta = Trim(txtCodigo(3).Text)
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{" & Tabla & ".fecfactu}"
        TipCod = "F"
        If Not PonerDesdeHasta(cDesde, cHasta, "", "", "pDHfechaFac= """) Then Exit Sub
    End If
    
    'D/H Colectivo
    cDesde = Trim(txtCodigo(4).Text)
    cHasta = Trim(txtCodigo(5).Text)
    nDesde = txtNombre(4).Text
    nHasta = txtNombre(5).Text
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{ssocio.codcoope}"
        TipCod = "N"
        If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHColec= """) Then Exit Sub
    End If
    
    Dim AnoDesde As String
    Dim AnoHasta As String
    If Option1(1).Value Then
        If txtCodigo(2).Text = "" Or txtCodigo(3).Text = "" Then
            MsgBox "Debe introducir la fecha desde/hasta. Revise.", vbExclamation
            Exit Sub
        Else
            AnoDesde = Mid(txtCodigo(2).Text, 7, 4)
            AnoHasta = Mid(txtCodigo(3).Text, 7, 4)
            If AnoDesde <> AnoHasta Then
                MsgBox "El rango de fechas debe de estar dentro del año natural. Revise.", vbExclamation
                Exit Sub
            Else
                CargarPeriodos
            End If
        End If
    End If
    
    
    'Comprobar si hay registros a Mostrar antes de abrir el Informe
    cadTABLA = Tabla & " INNER JOIN ssocio ON " & Tabla & ".codsocio=ssocio.codsocio "
    
    
    If Option1(0) = True Then
        If HayRegParaInforme(cadTABLA, cadSelect) Then
           If Option1(0) = True Then
              cadTitulo = "Detalle Ventas por Cliente"
              '[Monica]28/02/2014
              If NumCod = 0 Then
                    cadNombreRPT = "rEstCliimp.rpt"
              Else
                    cadNombreRPT = "rEstCliAjena.rpt"
              End If
              LlamarImprimir
              'AbrirVisReport
           End If
        End If
    Else
        If HayRegistros(cadTABLA, cadSelect) Then
            If CargarTablaIntermedia(cadTABLA, cadSelect) Then
                cadTitulo = "Resumen Ventas por Cliente/Trimestre"
                If NumCod = 0 Then
                    cadNombreRPT = "rEstCliimp1.rpt"
                Else
                    cadNombreRPT = "rEstCliAjena1.rpt"
                End If
                cadParam = cadParam & "Importe= " & DBSet(txtCodigo(6).Text, "N") & "|"
                numParam = numParam + 1
                cadFormula = "{tmpinformes.codusu} = " & vSesion.Codigo
                LlamarImprimir
            End If
        End If
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub CargarPeriodos()
Dim MesDesde As Integer
Dim MesHasta As Integer
Dim Periodos As Integer
Dim Anyo As String

    Anyo = Mid(txtCodigo(2).Text, 7, 4)

    If Mid(txtCodigo(3).Text, 4, 2) >= 1 Then Periodo1 = "between '" & Anyo & "-" & "01-01' and '" & Anyo & "-03-31'"
    If Mid(txtCodigo(3).Text, 4, 2) >= 3 Then Periodo2 = "between '" & Anyo & "-" & "04-01' and '" & Anyo & "-06-30'"
    If Mid(txtCodigo(3).Text, 4, 2) >= 6 Then Periodo3 = "between '" & Anyo & "-" & "07-01' and '" & Anyo & "-09-30'"
    If Mid(txtCodigo(3).Text, 4, 2) >= 9 Then Periodo4 = "between '" & Anyo & "-" & "10-01' and '" & Anyo & "-12-31'"

    cadParam = cadParam & "pPeriodo1=""01/01/" & Anyo & " - 31/03/" & Anyo & """|"
    cadParam = cadParam & "pPeriodo2=""01/04/" & Anyo & " - 30/06/" & Anyo & """|"
    cadParam = cadParam & "pPeriodo3=""01/07/" & Anyo & " - 30/09/" & Anyo & """|"
    cadParam = cadParam & "pPeriodo4=""01/10/" & Anyo & " - 31/12/" & Anyo & """|"
    numParam = numParam + 4

End Sub



Private Function CargarTablaIntermedia(Tabla As String, vSelect As String) As Boolean
Dim Sql As String
Dim Sql2 As String
Dim SqlAux As String
Dim Rs As ADODB.Recordset
Dim Rs2 As ADODB.Recordset
Dim CadValues As String
Dim sTabla As String

    On Error GoTo eCargarTablaIntermedia
    
    CargarTablaIntermedia = False
    
    Sql = "delete from tmpinformes where codusu = " & vSesion.Codigo
    Conn.Execute Sql
                                            'socio, periodo, baseimpo,cuotaiva,totalfac, nrofacturas
    Sql = "insert into tmpinformes (codusu, codigo1, campo1, importe1, importe2, importe3, importe4) values "
    
    If NumCod = 0 Then
        sTabla = "schfac"
    Else
        sTabla = "schfacr"
    End If
    
    
    SqlAux = "select " & sTabla & ".codsocio, sum(totalfac) from " & Tabla & " where (1=1) "
    If vSelect <> "" Then SqlAux = SqlAux & " and  " & vSelect
    SqlAux = SqlAux & " group by 1 "
    
    If txtCodigo(6).Text <> "" Then
        SqlAux = SqlAux & " having sum(totalfac) > " & DBSet(txtCodigo(6).Text, "N")
    End If
    
    
    
    CadValues = ""
    
    Set Rs = New ADODB.Recordset
    Rs.Open SqlAux, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    While Not Rs.EOF
        ' periodo 1
        Sql2 = "select " & sTabla & ".codsocio, sum(coalesce(baseimp1,0)+ coalesce(baseimp2,0)+ coalesce(baseimp3,0)) base,    "
        Sql2 = Sql2 & " sum(coalesce(impoiva1,0)+ coalesce(impoiva2,0)+ coalesce(impoiva3,0)) iva, sum(totalfac) total, "
        Sql2 = Sql2 & " count(*) nfactu "
        Sql2 = Sql2 & " from " & Tabla
        Sql2 = Sql2 & " where " & sTabla & ".codsocio = " & DBSet(Rs!codsocio, "N")
        If vSelect <> "" Then Sql2 = Sql2 & " and  " & vSelect
        Sql2 = Sql2 & " and " & sTabla & ".fecfactu " & Periodo1
        Sql2 = Sql2 & " group by 1"
        
        Set Rs2 = New ADODB.Recordset
        Rs2.Open Sql2, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
        If Not Rs2.EOF Then
            CadValues = CadValues & "(" & vSesion.Codigo & "," & DBSet(Rs!codsocio, "N") & ",1," & DBSet(Rs2!base, "N") & ","
            CadValues = CadValues & DBSet(Rs2!iva, "N") & "," & DBSet(Rs2!Total, "N") & "," & DBSet(Rs2!nfactu, "N") & "),"
        End If
        
        Set Rs2 = Nothing
    
        ' periodo 2
        Sql2 = "select " & sTabla & ".codsocio, sum(coalesce(baseimp1,0)+ coalesce(baseimp2,0)+ coalesce(baseimp3,0)) base,    "
        Sql2 = Sql2 & " sum(coalesce(impoiva1,0)+ coalesce(impoiva2,0)+ coalesce(impoiva3,0)) iva, sum(totalfac) total, "
        Sql2 = Sql2 & " count(*) nfactu "
        Sql2 = Sql2 & " from " & Tabla
        Sql2 = Sql2 & " where " & sTabla & ".codsocio = " & DBSet(Rs!codsocio, "N")
        If vSelect <> "" Then Sql2 = Sql2 & " and  " & vSelect
        Sql2 = Sql2 & " and " & sTabla & ".fecfactu " & Periodo2
        Sql2 = Sql2 & " group by 1"
        
        Set Rs2 = New ADODB.Recordset
        Rs2.Open Sql2, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
        If Not Rs2.EOF Then
            CadValues = CadValues & "(" & vSesion.Codigo & "," & DBSet(Rs!codsocio, "N") & ",2," & DBSet(Rs2!base, "N") & ","
            CadValues = CadValues & DBSet(Rs2!iva, "N") & "," & DBSet(Rs2!Total, "N") & "," & DBSet(Rs2!nfactu, "N") & "),"
        End If
        
        Set Rs2 = Nothing
    
        ' periodo 3
        Sql2 = "select " & sTabla & ".codsocio, sum(coalesce(baseimp1,0)+ coalesce(baseimp2,0)+ coalesce(baseimp3,0)) base,    "
        Sql2 = Sql2 & " sum(coalesce(impoiva1,0)+ coalesce(impoiva2,0)+ coalesce(impoiva3,0)) iva, sum(totalfac) total, "
        Sql2 = Sql2 & " count(*) nfactu "
        Sql2 = Sql2 & " from " & Tabla
        Sql2 = Sql2 & " where " & sTabla & ".codsocio = " & DBSet(Rs!codsocio, "N")
        If vSelect <> "" Then Sql2 = Sql2 & " and  " & vSelect
        Sql2 = Sql2 & " and " & sTabla & ".fecfactu " & Periodo3
        Sql2 = Sql2 & " group by 1"
        
        Set Rs2 = New ADODB.Recordset
        Rs2.Open Sql2, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
        If Not Rs2.EOF Then
            CadValues = CadValues & "(" & vSesion.Codigo & "," & DBSet(Rs!codsocio, "N") & ",3," & DBSet(Rs2!base, "N") & ","
            CadValues = CadValues & DBSet(Rs2!iva, "N") & "," & DBSet(Rs2!Total, "N") & "," & DBSet(Rs2!nfactu, "N") & "),"
        End If
        
        Set Rs2 = Nothing
    
        ' periodo 4
        Sql2 = "select " & sTabla & ".codsocio, sum(coalesce(baseimp1,0)+ coalesce(baseimp2,0)+ coalesce(baseimp3,0)) base,    "
        Sql2 = Sql2 & " sum(coalesce(impoiva1,0)+ coalesce(impoiva2,0)+ coalesce(impoiva3,0)) iva, sum(totalfac) total, "
        Sql2 = Sql2 & " count(*) nfactu "
        Sql2 = Sql2 & " from " & Tabla
        Sql2 = Sql2 & " where " & sTabla & ".codsocio = " & DBSet(Rs!codsocio, "N")
        If vSelect <> "" Then Sql2 = Sql2 & " and  " & vSelect
        Sql2 = Sql2 & " and " & sTabla & ".fecfactu " & Periodo4
        Sql2 = Sql2 & " group by 1"
        
        Set Rs2 = New ADODB.Recordset
        Rs2.Open Sql2, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
        If Not Rs2.EOF Then
            CadValues = CadValues & "(" & vSesion.Codigo & "," & DBSet(Rs!codsocio, "N") & ",4," & DBSet(Rs2!base, "N") & ","
            CadValues = CadValues & DBSet(Rs2!iva, "N") & "," & DBSet(Rs2!Total, "N") & "," & DBSet(Rs2!nfactu, "N") & "),"
        End If
        
        Set Rs2 = Nothing
    
        Rs.MoveNext
    Wend
    
    If CadValues <> "" Then
        CadValues = Mid(CadValues, 1, Len(CadValues) - 1)
        Conn.Execute Sql & CadValues
    End If
    
    
    Set Rs = Nothing
    
    CargarTablaIntermedia = True
    Exit Function
    
    
eCargarTablaIntermedia:
    MuestraError Err.Number, "Cargar Tabla Intermedia", Err.Description
End Function


Private Function CargarTablaIntermediaNew(Tabla As String, vSelect As String) As Boolean
Dim Sql As String
Dim Sql2 As String
Dim SqlAux As String
Dim Rs As ADODB.Recordset
Dim Rs2 As ADODB.Recordset
Dim CadValues As String
Dim sTabla As String

    On Error GoTo eCargarTablaIntermedia
    
    CargarTablaIntermediaNew = False
    
    Sql = "delete from tmpinformes where codusu = " & vSesion.Codigo
    Conn.Execute Sql
                                            'socio, total, total1, total2, total3, total4, nfact, nfac1, nfac2, nfac3, nfac4
    Sql = "insert into tmpinformes (codusu, codigo1, importe1, importeb1, importe2, importeb2, importe3, importeb3, importe4, importeb4, importe5, importeb5) values "
    
    
    If NumCod = 0 Then
        sTabla = "schfac"
    Else
        sTabla = "schfacr"
    End If
    
    SqlAux = "select " & sTabla & ".codsocio, sum(totalfac) total , count(*) nfac from " & Tabla & " where (1=1) "
    If vSelect <> "" Then SqlAux = SqlAux & " and  " & vSelect
    SqlAux = SqlAux & " group by 1 "
    
    If txtCodigo(6).Text <> "" Then
        SqlAux = SqlAux & " having sum(totalfac) > " & DBSet(txtCodigo(6).Text, "N")
    End If
    
    CadValues = ""
    
    Set Rs = New ADODB.Recordset
    Rs.Open SqlAux, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    While Not Rs.EOF
        CadValues = CadValues & "(" & vSesion.Codigo & "," & DBSet(Rs!codsocio, "N") & "," & DBSet(Rs!Total, "N") & "," & DBSet(Rs!nfac, "N") & ","
        ' periodo 1
        Sql2 = "select " & sTabla & ".codsocio, "
        Sql2 = Sql2 & " sum(totalfac) total, "
        Sql2 = Sql2 & " count(*) nfactu "
        Sql2 = Sql2 & " from " & Tabla
        Sql2 = Sql2 & " where " & sTabla & ".codsocio = " & DBSet(Rs!codsocio, "N")
        If vSelect <> "" Then Sql2 = Sql2 & " and  " & vSelect
        Sql2 = Sql2 & " and " & sTabla & ".fecfactu " & Periodo1
        Sql2 = Sql2 & " group by 1"
        
        Set Rs2 = New ADODB.Recordset
        Rs2.Open Sql2, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
        If Not Rs2.EOF Then
            CadValues = CadValues & DBSet(Rs2!Total, "N") & "," & DBSet(Rs2!nfactu, "N") & ","
        Else
            CadValues = CadValues & "0,0,"
        End If
        
        Set Rs2 = Nothing
    
        ' periodo 2
        Sql2 = "select " & sTabla & ".codsocio, "
        Sql2 = Sql2 & " sum(totalfac) total, "
        Sql2 = Sql2 & " count(*) nfactu "
        Sql2 = Sql2 & " from " & Tabla
        Sql2 = Sql2 & " where " & sTabla & ".codsocio = " & DBSet(Rs!codsocio, "N")
        If vSelect <> "" Then Sql2 = Sql2 & " and  " & vSelect
        Sql2 = Sql2 & " and " & sTabla & ".fecfactu " & Periodo2
        Sql2 = Sql2 & " group by 1"
        
        Set Rs2 = New ADODB.Recordset
        Rs2.Open Sql2, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
        If Not Rs2.EOF Then
            CadValues = CadValues & DBSet(Rs2!Total, "N") & "," & DBSet(Rs2!nfactu, "N") & ","
        Else
            CadValues = CadValues & "0,0,"
        End If
        
        Set Rs2 = Nothing
    
        ' periodo 3
        Sql2 = "select " & sTabla & ".codsocio,    "
        Sql2 = Sql2 & " sum(totalfac) total, "
        Sql2 = Sql2 & " count(*) nfactu "
        Sql2 = Sql2 & " from " & Tabla
        Sql2 = Sql2 & " where " & sTabla & ".codsocio = " & DBSet(Rs!codsocio, "N")
        If vSelect <> "" Then Sql2 = Sql2 & " and  " & vSelect
        Sql2 = Sql2 & " and " & sTabla & ".fecfactu " & Periodo3
        Sql2 = Sql2 & " group by 1"
        
        Set Rs2 = New ADODB.Recordset
        Rs2.Open Sql2, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
        If Not Rs2.EOF Then
            CadValues = CadValues & DBSet(Rs2!Total, "N") & "," & DBSet(Rs2!nfactu, "N") & ","
        Else
            CadValues = CadValues & "0,0,"
        End If
        
        Set Rs2 = Nothing
    
        ' periodo 4
        Sql2 = "select " & sTabla & ".codsocio,    "
        Sql2 = Sql2 & "  sum(totalfac) total, "
        Sql2 = Sql2 & " count(*) nfactu "
        Sql2 = Sql2 & " from " & Tabla
        Sql2 = Sql2 & " where " & sTabla & ".codsocio = " & DBSet(Rs!codsocio, "N")
        If vSelect <> "" Then Sql2 = Sql2 & " and  " & vSelect
        Sql2 = Sql2 & " and " & sTabla & ".fecfactu " & Periodo4
        Sql2 = Sql2 & " group by 1"
        
        Set Rs2 = New ADODB.Recordset
        Rs2.Open Sql2, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
        If Not Rs2.EOF Then
            CadValues = CadValues & DBSet(Rs2!Total, "N") & "," & DBSet(Rs2!nfactu, "N") & "),"
        Else
            CadValues = CadValues & "0,0),"
        End If
        
        Set Rs2 = Nothing
    
        Rs.MoveNext
    Wend
    
    If CadValues <> "" Then
        CadValues = Mid(CadValues, 1, Len(CadValues) - 1)
        Conn.Execute Sql & CadValues
    End If
    
    
    Set Rs = Nothing
    
    CargarTablaIntermediaNew = True
    Exit Function
    
    
eCargarTablaIntermedia:
    MuestraError Err.Number, "Cargar Tabla Intermedia", Err.Description
End Function





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
    limpiar Me

    'IMAGES para busqueda
     Me.imgBuscar(0).Picture = frmPpal.imgListImages16.ListImages(1).Picture
     Me.imgBuscar(1).Picture = frmPpal.imgListImages16.ListImages(1).Picture
     Me.imgBuscar(4).Picture = frmPpal.imgListImages16.ListImages(1).Picture
     Me.imgBuscar(5).Picture = frmPpal.imgListImages16.ListImages(1).Picture

    '###Descomentar
'    CommitConexion
         
    FrameCobrosVisible True, h, w
    indFrame = 5
    If NumCod = 0 Then
        Tabla = "schfac"
    Else
        Tabla = "schfacr"
    End If
    
    Frame1.visible = False
    
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

Private Sub frmCol_DatoSeleccionado(CadenaSeleccion As String)
'Form de Consulta de Colectivos
    txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "00")
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
        
        Case 4, 5 'COLECTIVO
            AbrirFrmColectivo (Index)
        
    End Select
    PonerFoco txtCodigo(indCodigo)
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

Private Sub Option1_Click(Index As Integer)
Dim b As Boolean
    b = (Option1(1).Value = True)
    Frame1.visible = b
    If b Then PonerFoco txtCodigo(6)
End Sub

Private Sub txtCodigo_GotFocus(Index As Integer)
    ConseguirFoco txtCodigo(Index), 3
End Sub

Private Sub txtCodigo_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    KEYdown KeyCode
End Sub

Private Sub txtCodigo_KeyPress(Index As Integer, KeyAscii As Integer)
'14/02/2007 antes
'    KEYpress KeyAscii
    If KeyAscii = teclaBuscar Then
        Select Case Index
            Case 0: KEYBusqueda KeyAscii, 0 'cliente desde
            Case 1: KEYBusqueda KeyAscii, 1 'cliente hasta
            Case 4: KEYBusqueda KeyAscii, 4 'colectivo desde
            Case 5: KEYBusqueda KeyAscii, 5 'colectivo hasta
            Case 2: KEYFecha KeyAscii, 2 'fecha desde
            Case 3: KEYFecha KeyAscii, 3 'fecha hasta
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
        
        Case 2, 3 'FECHAS
            If txtCodigo(Index).Text <> "" Then PonerFormatoFecha txtCodigo(Index)
            
        Case 4, 5 'COLECTIVO
            txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), "scoope", "nomcoope", "codcoope", "N")
            If txtCodigo(Index).Text <> "" Then txtCodigo(Index).Text = Format(txtCodigo(Index).Text, "00")
            
        Case 6 ' IMPORTE DESDE EN CASO DE QUE SEA RESUMEN
            If txtCodigo(Index).Text <> "" Then PonerFormatoDecimal txtCodigo(Index), 3
                        
    End Select
End Sub

Private Sub FrameCobrosVisible(visible As Boolean, ByRef h As Integer, ByRef w As Integer)
    Me.FrameCobros.visible = visible
    If visible = True Then
        Me.FrameCobros.Top = -90
        Me.FrameCobros.Left = 0
        Me.FrameCobros.Height = 6015
        Me.FrameCobros.Width = 6555
        w = Me.FrameCobros.Width
        h = Me.FrameCobros.Height
    End If
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

Private Sub LlamarImprimir()
    With frmImprimir
        .FormulaSeleccion = cadFormula
        .OtrosParametros = cadParam
        .NumeroParametros = numParam
        .SoloImprimir = False
        .Titulo = cadTitulo
        .NombreRPT = cadNombreRPT
        .Opcion = 0
        .Show vbModal
    End With
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
 
Private Sub AbrirVisReport()
    Screen.MousePointer = vbHourglass
    CadenaDesdeOtroForm = ""
    With frmVisReport
        .FormulaSeleccion = cadFormula
        '.SoloImprimir = (Me.OptVisualizar(indFrame).Value = 1)
        .OtrosParametros = cadParam
        .NumeroParametros = numParam
        '##descomen
'        .MostrarTree = MostrarTree
'        .Informe = MIPATH & Nombre
'        .InfConta = InfConta
        '##
        
'        If NombreSubRptConta <> "" Then
'            .SubInformeConta = NombreSubRptConta
'        Else
'            .SubInformeConta = ""
'        End If
        '##descomen
'        .ConSubInforme = ConSubInforme
        '##
        .Opcion = ""
'        .ExportarPDF = (chkEMAIL.Value = 1)
        .Show vbModal
    End With
    
'    If Me.chkEMAIL.Value = 1 Then
'    '####Descomentar
'        If CadenaDesdeOtroForm <> "" Then frmEMail.Show vbModal
'    End If
    Unload Me
End Sub

Private Sub AbrirEMail()
    If CadenaDesdeOtroForm <> "" Then frmEMail.Show vbModal
End Sub


Private Function HayRegistros(cTabla As String, cWhere As String) As Boolean
'Comprobar si hay registros a Mostrar antes de abrir el Informe
Dim Sql As String
Dim Rs As ADODB.Recordset
Dim sTabla As String


    If NumCod = 0 Then
        sTabla = "schfac"
    Else
        sTabla = "schfacr"
    End If

    Sql = "Select " & sTabla & ".codsocio, sum(totalfac) FROM " & QuitarCaracterACadena(cTabla, "_1")
    If cWhere <> "" Then
        cWhere = QuitarCaracterACadena(cWhere, "{")
        cWhere = QuitarCaracterACadena(cWhere, "}")
        cWhere = QuitarCaracterACadena(cWhere, "_1")
        Sql = Sql & " WHERE " & cWhere
    End If
    Sql = Sql & " group by 1 "
    Sql = Sql & " having sum(totalfac) > " & DBSet(txtCodigo(6).Text, "N")
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    If Rs.EOF Then
        MsgBox "No hay datos para mostrar en el Informe.", vbInformation
        HayRegistros = False
    Else
        HayRegistros = True
    End If

End Function

