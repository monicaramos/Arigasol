VERSION 5.00
Begin VB.Form frmExpTPV2 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   8145
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7290
   Icon            =   "frmExpTPV2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8145
   ScaleWidth      =   7290
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtcodigo 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   14
      Left            =   1695
      MaxLength       =   10
      TabIndex        =   2
      Tag             =   "Código Postal|T|S|||clientes|codposta|||"
      Top             =   2160
      Width           =   1050
   End
   Begin VB.TextBox txtcodigo 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   15
      Left            =   1695
      MaxLength       =   10
      TabIndex        =   3
      Tag             =   "Código Postal|T|S|||clientes|codposta|||"
      Top             =   2520
      Width           =   1050
   End
   Begin VB.TextBox txtNombre 
      BackColor       =   &H80000018&
      Height          =   285
      Index           =   1
      Left            =   2730
      Locked          =   -1  'True
      TabIndex        =   24
      Top             =   1545
      Width           =   4065
   End
   Begin VB.TextBox txtNombre 
      BackColor       =   &H80000018&
      Height          =   285
      Index           =   0
      Left            =   2730
      Locked          =   -1  'True
      TabIndex        =   23
      Top             =   1170
      Width           =   4065
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   0
      Left            =   480
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   4050
      Width           =   6255
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   1
      Left            =   480
      TabIndex        =   6
      Text            =   "Text1"
      Top             =   4650
      Width           =   6255
   End
   Begin VB.CommandButton cmdExport 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   3690
      TabIndex        =   10
      Top             =   7500
      Width           =   1425
   End
   Begin VB.CommandButton cmdSal 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   5310
      TabIndex        =   11
      Top             =   7500
      Width           =   1425
   End
   Begin VB.TextBox txtcodigo 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   0
      Left            =   1710
      MaxLength       =   6
      TabIndex        =   0
      Top             =   1170
      Width           =   945
   End
   Begin VB.TextBox txtcodigo 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   1
      Left            =   1710
      MaxLength       =   6
      TabIndex        =   1
      Top             =   1530
      Width           =   945
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "frmExpTPV2.frx":000C
      Left            =   1650
      List            =   "frmExpTPV2.frx":000E
      TabIndex        =   4
      Text            =   "Combo1"
      Top             =   3120
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   2
      Left            =   480
      TabIndex        =   7
      Text            =   "Text1"
      Top             =   5250
      Width           =   6255
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   3
      Left            =   480
      TabIndex        =   8
      Text            =   "Text1"
      Top             =   5850
      Width           =   6255
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   4
      Left            =   480
      TabIndex        =   9
      Text            =   "Text1"
      Top             =   6570
      Width           =   6255
   End
   Begin VB.Label Label4 
      Caption         =   "Fecha Alta"
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
      Height          =   255
      Index           =   16
      Left            =   540
      TabIndex        =   27
      Top             =   1950
      Width           =   795
   End
   Begin VB.Label Label4 
      Caption         =   "Desde"
      Height          =   195
      Index           =   15
      Left            =   810
      TabIndex        =   26
      Top             =   2190
      Width           =   465
   End
   Begin VB.Label Label4 
      Caption         =   "Hasta"
      Height          =   195
      Index           =   14
      Left            =   810
      TabIndex        =   25
      Top             =   2550
      Width           =   420
   End
   Begin VB.Image imgFec 
      Height          =   240
      Index           =   14
      Left            =   1380
      Picture         =   "frmExpTPV2.frx":0010
      ToolTipText     =   "Buscar fecha"
      Top             =   2160
      Width           =   240
   End
   Begin VB.Image imgFec 
      Height          =   240
      Index           =   15
      Left            =   1380
      Picture         =   "frmExpTPV2.frx":009B
      ToolTipText     =   "Buscar fecha"
      Top             =   2520
      Width           =   240
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
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
      ForeColor       =   &H00972E0B&
      Height          =   195
      Index           =   21
      Left            =   570
      TabIndex        =   22
      Top             =   840
      Width           =   525
   End
   Begin VB.Label Label6 
      Caption         =   "Exportación para TPV"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   345
      Left            =   600
      TabIndex        =   21
      Top             =   210
      Width           =   5145
   End
   Begin VB.Label Label1 
      Caption         =   "Fichero para exportación (CLIENTES)"
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
      Height          =   255
      Index           =   0
      Left            =   480
      TabIndex        =   20
      Top             =   3810
      Width           =   6255
   End
   Begin VB.Label Label1 
      Caption         =   "Fichero para exportación (CLIENTESBASE)"
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
      Height          =   255
      Index           =   1
      Left            =   480
      TabIndex        =   19
      Top             =   4350
      Width           =   6255
   End
   Begin VB.Label lblInf 
      Height          =   225
      Left            =   510
      TabIndex        =   18
      Top             =   7020
      Width           =   6255
   End
   Begin VB.Label lblTar 
      Caption         =   "Desde"
      Height          =   255
      Index           =   0
      Left            =   840
      TabIndex        =   17
      Top             =   1200
      Width           =   2055
   End
   Begin VB.Label lblTar 
      Caption         =   "Hasta"
      Height          =   255
      Index           =   1
      Left            =   840
      TabIndex        =   16
      Top             =   1560
      Width           =   705
   End
   Begin VB.Label lblTar 
      Caption         =   "Estado"
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
      Height          =   255
      Index           =   2
      Left            =   510
      TabIndex        =   15
      Top             =   3150
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Fichero para exportación (VEHICULOS)"
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
      Height          =   255
      Index           =   2
      Left            =   480
      TabIndex        =   14
      Top             =   5610
      Width           =   6255
   End
   Begin VB.Label Label1 
      Caption         =   "Fichero para exportación (CODIGOSCLIENTE)"
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
      Height          =   255
      Index           =   3
      Left            =   480
      TabIndex        =   13
      Top             =   5010
      Width           =   6255
   End
   Begin VB.Label Label1 
      Caption         =   "Fichero para exportación (PROMOCLIENT)"
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
      Height          =   255
      Index           =   4
      Left            =   480
      TabIndex        =   12
      Top             =   6270
      Width           =   6255
   End
End
Attribute VB_Name = "frmExpTPV2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents frmC As frmCal 'calendario fecha
Attribute frmC.VB_VarHelpID = -1



Dim sql As String
Dim rs As ADODB.Recordset
Dim Rs2 As ADODB.Recordset
Dim NF1 As Integer
Dim NF2 As Integer
Dim NF3 As Integer
Dim NF4 As Integer
Dim NF5 As Integer
Dim Registro As String
Dim mAux As String
'-- Variables de apoyo para montar luego registros
Dim TARJETA As String
Dim Tipo As String
Dim PRODUCTO As String
Dim CLIENTE As String
Dim ACTIVA As String
Dim RIESGO As String
Dim IMPORTE As String
Dim DESCUENTO As String
Dim PIN As String
'---
Dim MATRICULA As String
Dim Nombre As String
Dim NIF As String
Dim DIRECCION As String
Dim POBLACION As String
Dim M As String
Dim DesdeTar As String
Dim HastaTar As String
'---
Dim TABULADOR As String

Dim indice As Integer
Dim Modo As Byte




Private Sub cmdExport_Click()
    '-- Controlamos la imputación de tarjetas
    DesdeTar = txtCodigo(0)
    HastaTar = txtCodigo(1)
    If HastaTar < DesdeTar Then
        MsgBox "El valor desde tarjeta no puede ser mayor que hasta tarjeta", vbInformation
        Exit Sub
    End If
    
    If txtCodigo(14).Text <> "" And txtCodigo(15).Text <> "" Then
        If CDate(txtCodigo(15).Text) < CDate(txtCodigo(14).Text) Then
            MsgBox "El valor desde fecha no puede ser mayor que hasta fecha", vbInformation
            Exit Sub
        End If
     End If
     
    '-- Exportando asociados
    If Exp_Tarjetas(DesdeTar, HastaTar) Then
        MsgBox "Proceso realizado correctamente.", vbExclamation
        cmdSal_Click
    End If
    
End Sub

Private Sub cmdSal_Click()
    Unload Me
End Sub

Private Sub Form_Load()


    TABULADOR = Chr(9)
    
    If Dir(App.path & "\Exportaciones", vbDirectory) = "" Then MkDir App.path & "\Exportaciones"
    
    Text1(0) = App.path & "\Exportaciones\CLIENTES.TXT"
    Text1(1) = App.path & "\Exportaciones\CLIENTESBASE.TXT"
    Text1(2) = App.path & "\Exportaciones\CODIGOSCLIENTE.TXT"
    Text1(3) = App.path & "\Exportaciones\VEHICULOS.TXT"
    Text1(4) = App.path & "\Exportaciones\PROMOCLIENT.TXT"
    
    txtCodigo(14).Text = Format(Now, "dd/mm/yyyy")
    txtCodigo(15).Text = Format(Now, "dd/mm/yyyy")
    
    
    
    Combo1.Clear
    Combo1.AddItem "Todas"
    Combo1.ItemData(Combo1.NewIndex) = 0
    Combo1.AddItem "Inactivas"
    Combo1.ItemData(Combo1.NewIndex) = 1
    Combo1.AddItem "Activas"
    Combo1.ItemData(Combo1.NewIndex) = 2
    Combo1.ListIndex = 0
    
End Sub

Private Function Exp_Tarjetas(dTar As String, hTar As String) As Boolean
    Dim Estado As Integer
    Dim Procesar As Boolean
    Dim Entidad As String
    Dim Nombre As String
    Dim Provincia As String
    Dim CodPostal As String
    Dim i, i2, NumLin As Long
    Dim L1 As String ' Linea de grabación para fichero 1
    Dim L2 As String ' Linea de grabación para fichero 2
    Dim L3 As String ' Linea de grabación para fichero 3
    Dim L4 As String ' Linea de grabación para fichero 4
    Dim L5 As String ' Linea de grabación para fichero 5
    
    On Error GoTo eExp_Tarjetas
    
    Exp_Tarjetas = False
    
    
    
    Estado = Combo1.ListIndex
    NF1 = FreeFile()
    Open Text1(0) For Output As #NF1
    NF2 = FreeFile()
    Open Text1(1) For Output As #NF2
    NF3 = FreeFile()
    Open Text1(2) For Output As #NF3
    NF4 = FreeFile()
    Open Text1(3) For Output As #NF4
    NF5 = FreeFile()
    Open Text1(4) For Output As #NF5
    sql = "SELECT * FROM starje WHERE (1=1) "
    If dTar <> "" Then sql = sql & " and numtarje >= " & DBSet(dTar, "N")
    If hTar <> "" Then sql = sql & " AND numtarje <= " & DBSet(hTar, "N")
    
    '[Monica]29/05/2014: añadida la fecha de alta de tarjetas
    If txtCodigo(14).Text <> "" Then sql = sql & " and fecalta >= " & DBSet(txtCodigo(14).Text, "F")
    If txtCodigo(15).Text <> "" Then sql = sql & " and fecalta <= " & DBSet(txtCodigo(15).Text, "F")
    
    Select Case Combo1.ListIndex
        Case 0 'todas
        Case 1 'inactivas
            sql = sql & " and (Estado = 1 Or Estado = 4)"
        Case 2 'activas
            sql = sql & " and not (Estado = 1 Or Estado = 4)"
    End Select
    
    Set rs = New ADODB.Recordset
    rs.Open sql, Conn, , , adCmdText
    If rs.EOF Then
        MsgBox "No hay tarjetas a exportar entre esos límites", vbExclamation
        
        rs.Close
        
        Close #NF1
        Close #NF2
        Close #NF3
        Close #NF4
        Close #NF5
    
        Exit Function
    Else
        '-- Grabación de las líneas de cabecera
        L1 = "CODIGOCLIENTE" & TABULADOR ' codigocliente
        L1 = L1 & "DESCRIPCION" & TABULADOR ' descripcion
        L1 = L1 & "CAE" & TABULADOR ' cae
        L1 = L1 & "NOMBRE" & TABULADOR ' nombre
        L1 = L1 & "NIF" & TABULADOR ' nif
        L1 = L1 & "DIRECCION" & TABULADOR ' direccion
        L1 = L1 & "CP" & TABULADOR ' cp
        L1 = L1 & "POBLACION" & TABULADOR ' poblacion
        L1 = L1 & "PROVINCIA" & TABULADOR ' provincia
        L1 = L1 & "IDCOMUNIDAD" & TABULADOR ' idcomunidad
        L1 = L1 & "PAIS" & TABULADOR ' pais
        L1 = L1 & "IDIOMA" & TABULADOR ' idioma
        L1 = L1 & "CREDITO" & TABULADOR ' credito
        L1 = L1 & "PRODUCTESAUT"  ' productesaut (Normal)
        '-- L2
        L2 = "IDBASE" & TABULADOR ' idbase
        L2 = L2 & "IDCLIENTE" & TABULADOR ' idcliente
        L2 = L2 & "ACTIVO" ' activo
        '-- L3
        L3 = "IDCODIGOALT" & TABULADOR ' idcodigoalt
        L3 = L3 & "TIPOCODIGO" & TABULADOR ' tipocodigo
        L3 = L3 & "IDCLIENTE" & TABULADOR ' idcliente
        L3 = L3 & "IDVEHICULO" & TABULADOR ' idvehiculo
        L3 = L3 & "TIPOTARJETA" ' tipotarjeta
        '-- L4
        L4 = "IDVEHICULO" & TABULADOR ' idvehiculo
        L4 = L4 & "MATRICULA" & TABULADOR 'matricula
        L4 = L4 & "DESCRIPCION" & TABULADOR ' descricpcion
        L4 = L4 & "IDCLIENTE" & TABULADOR ' idcliente
        L4 = L4 & "NUMVEHICULO" & TABULADOR ' numvehiculo
        L4 = L4 & "PRODUCTESAUT"  ' productesaut (Normal)
        '-- L5
        L5 = "IDPROMOCION" & TABULADOR ' idpromocion
        L5 = L5 & "IDCLIENTE" & TABULADOR 'idcliente
        L5 = L5 & "IDBASE" ' idbase
        '-- Y ahora a grabar la información
        If L1 <> "" Then Print #NF1, L1
        If L2 <> "" Then Print #NF2, L2
        If L3 <> "" Then Print #NF3, L3
        If L4 <> "" Then Print #NF4, L4
        If L5 <> "" Then Print #NF5, L5
        rs.MoveFirst
        i2 = rs.RecordCount
        While Not rs.EOF
            '-- Inicializamos las líneas de grabación
            L1 = "": L2 = "": L3 = "": L4 = "": L5 = ""
            '-- Vemos cuales se procesan en funcion de lo solicitado y del estado
            Procesar = False ' por defecto no se procesa
            Select Case Estado
                Case 0 ' se procesan todas
                    Procesar = True
                Case 1 ' solo inactivas
                    If rs!Estado = 1 Or rs!Estado = 4 Then Procesar = True
                Case 2 ' solo activas
                    If rs!Estado <> 1 And rs!Estado <> 4 Then Procesar = True
            End Select
            If Procesar Then
                i = i + 1
                NumLin = 0
                lblInf.Caption = "Procesando registro " & CStr(i)
                lblInf.Refresh
                '-- Montamos L1 --
                TARJETA = "9724000030" & Format(DBLet(rs!Numtarje), "000000")
                CLIENTE = Right(TARJETA, 5)
                MATRICULA = DBLet(rs!MATRICUL)
                PRODUCTO = "N" '--Se supone que carburanteantes era C [3.6.20]
'[Monica]29/05/2014: controlamos gasoleo b, que en nuestro caso es tiptarje
'                If rs!GasoleoB <> 0 Then
'                    PRODUCTO = "R" '--Ahora controla si hay Gasoleo B [3.6.20]
'                End If
                If rs!tiptarje = 1 Then
                    PRODUCTO = "R" '--Ahora controla si hay Gasoleo B [3.6.20]
                End If



                '-- Ahora en el lugar del descuento tiene que figurar la entidad"
                sql = "select * from ssocio where codsocio = " & CStr(rs!codsocio)
                Set Rs2 = New ADODB.Recordset
                Rs2.Open sql, Conn, , , adCmdText
                If Not Rs2.EOF Then
                    Entidad = CStr(Rs2!codcoope)
                Else
                    Entidad = ""
                End If
'[Monica]29/05/2014: tenemos todo el nombre junto en la starje
'                If (rs.Fields("Apellido1") = "" And rs.Fields("Apellido2") = "") Or (IsNull(rs.Fields("Apellido1")) And IsNull(rs.Fields("Apellido2"))) Then
'                    mAux = Trim(rs.Fields("Nombre"))
'                Else
'                    mAux = Trim(rs.Fields("Apellido1")) & " " & Trim(rs.Fields("Apellido2")) & ", " & Trim(rs.Fields("Nombre"))
'                End If
                mAux = DBLet(rs!nomtarje, "T")
                
                Nombre = DBLet(mAux)
'[Monica]29/05/2014: el nif es el del socio
                NIF = DBLet(Rs2!NIFsocio)
                
                sql = "select * from ssocio where codsocio = " & CStr(rs!codsocio)
                Set Rs2 = New ADODB.Recordset
                Rs2.Open sql, Conn, , , adCmdText
                If Not Rs2.EOF Then
                    DIRECCION = DBLet(Rs2!domsocio)
                    POBLACION = DBLet(Rs2!pobsocio)
                    CodPostal = DBLet(Rs2!CodPosta)
                    Provincia = DBLet(Rs2!Prosocio)
                Else
                    DIRECCION = ""
                    POBLACION = ""
                    CodPostal = ""
                    Provincia = ""
                End If
                '-- Montaje de LINEAS
                L1 = CLIENTE & TABULADOR ' codigocliente
                L1 = L1 & "" & TABULADOR ' descripcion
                L1 = L1 & "" & TABULADOR ' cae
                L1 = L1 & Nombre & TABULADOR ' nombre
                L1 = L1 & NIF & TABULADOR ' nif
                L1 = L1 & DIRECCION & TABULADOR ' direccion
                L1 = L1 & CodPostal & TABULADOR ' cp
                L1 = L1 & POBLACION & TABULADOR ' poblacion
                L1 = L1 & Provincia & TABULADOR ' provincia
                L1 = L1 & "1" & TABULADOR ' idcomunidad
                L1 = L1 & "011" & TABULADOR ' pais
                L1 = L1 & "1" & TABULADOR ' idioma
                L1 = L1 & "0" & TABULADOR ' credito
                If PRODUCTO = "R" Then
                    L1 = L1 & "2"  ' productesaut (Gasóleo B)
                Else
                    L1 = L1 & "16381"  ' productesaut (Normal)
                End If
                '-- L2
                L2 = "1" & TABULADOR ' idbase
                L2 = L2 & CLIENTE & TABULADOR ' cliente
                L2 = L2 & "1" ' activo
                '-- L3
                L3 = TARJETA & TABULADOR ' idcodigoalt
                L3 = L3 & "5" & TABULADOR ' tipocodigo
                L3 = L3 & CLIENTE & TABULADOR ' idcliente
                L3 = L3 & CLIENTE & TABULADOR ' idvehiculo
                L3 = L3 & "0" ' tipotarjeta
                '-- L4
                L4 = CLIENTE & TABULADOR ' idvehiculo
                L4 = L4 & MATRICULA & TABULADOR 'matricula
                L4 = L4 & "" & TABULADOR ' descricpcion
                L4 = L4 & CLIENTE & TABULADOR ' idcliente
                L4 = L4 & "1" & TABULADOR ' numvehiculo
                If PRODUCTO = "R" Then
                    L4 = L4 & "2"  ' productesaut (Gasóleo B)
                Else
                    L4 = L4 & "16381"  ' productesaut (Normal)
                End If
                '-- L5
                If Entidad <> "" And Entidad <> "0" And Entidad <> "6" Then
                    L5 = Entidad & TABULADOR ' idpromocion
                    L5 = L5 & CLIENTE & TABULADOR 'idcliente
                    L5 = L5 & "1" ' idbase
                End If
                '-- Y ahora a grabar la información
                If L1 <> "" Then Print #NF1, L1
                If L2 <> "" Then Print #NF2, L2
                If L3 <> "" Then Print #NF3, L3
                If L4 <> "" Then Print #NF4, L4
                If L5 <> "" Then Print #NF5, L5
            End If
            rs.MoveNext
        Wend
    End If
    rs.Close
    Close #NF1
    Close #NF2
    Close #NF3
    Close #NF4
    Close #NF5
    
    Exp_Tarjetas = True
    Exit Function
    
eExp_Tarjetas:
    Close #NF1
    Close #NF2
    Close #NF3
    Close #NF4
    Close #NF5


    MuestraError Err.Number, "Exportación de Tarjetas TPV", Err.Description
End Function



Private Function NumRegistros(Tabla As String) As Long
    Dim sql As String
    Dim rs As ADODB.Recordset
    sql = "Select Count(*) from " & Tabla
    NumRegistros = 0
    Set rs = New ADODB.Recordset
    rs.Open sql, Conn, , , adCmdText
    If Not rs.EOF Then
        If Not IsNull(rs.Fields(0)) Then
            NumRegistros = rs.Fields(0)
        End If
    End If
    rs.Close
End Function

Private Sub frmC_Selec(vFecha As Date)
    ' *** repasar si el camp es txtAux o Text1 ***
    txtCodigo(CByte(imgFec(15).Tag)).Text = Format(vFecha, "dd/mm/yyyy") '<===
    ' ********************************************
End Sub

Private Sub imgFec_Click(Index As Integer)
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

    frmC.Left = esq + imgFec(Index).Parent.Left + 30
    frmC.Top = dalt + imgFec(Index).Parent.Top + imgFec(Index).Height + menu - 40

    imgFec(15).Tag = Index '<===
    ' *** repasar si el camp es txtAux o Text1 ***
    If txtCodigo(Index).Text <> "" Then frmC.NovaData = txtCodigo(Index).Text
    ' ********************************************

    frmC.Show vbModal
    Set frmC = Nothing
    ' *** repasar si el camp es txtAux o Text1 ***
    PonerFoco txtCodigo(CByte(imgFec(15).Tag)) '<===
    ' ********************************************
End Sub

Private Sub txtCodigo_GotFocus(Index As Integer)
    indice = Index
    ConseguirFoco txtCodigo(Index), 3
End Sub

Private Sub txtCodigo_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    KEYdown KeyCode
End Sub

Private Sub txtCodigo_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = teclaBuscar Then
        If Modo = 1 Or Modo = 3 Or Modo = 4 Then
            Select Case Index
                Case 14: KEYFecha KeyAscii, 14 'desde fecha
                Case 15: KEYFecha KeyAscii, 15 'hasta fecha
           End Select
        End If
    Else
        KEYpress KeyAscii
    End If
End Sub

Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, Modo, cerrar
    If cerrar Then Unload Me
End Sub



Private Sub KEYFecha(KeyAscii As Integer, indice As Integer)
    KeyAscii = 0
    imgFec_Click (indice)
End Sub

Private Sub txtCodigo_LostFocus(Index As Integer)
Dim Nuevo As Boolean

    txtCodigo(Index).Text = Trim(txtCodigo(Index).Text)
    
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub
    
    Select Case Index
        Case 14, 15 ' desde hasta fecha
            PonerFormatoFecha txtCodigo(Index)
            
            
        Case 0, 1 ' tarjetas
            If txtCodigo(Index).Text <> "" Then txtCodigo(Index) = Format(txtCodigo(Index), "000000")
            txtNombre(Index) = PonerNombreDeCod(txtCodigo(Index), "starje", "nomtarje", "numtarje", "N")
    End Select
End Sub

