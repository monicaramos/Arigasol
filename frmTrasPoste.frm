VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmTrasPoste 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Traspaso Datos Poste"
   ClientHeight    =   4890
   ClientLeft      =   45
   ClientTop       =   2430
   ClientWidth     =   6825
   Icon            =   "frmTrasPoste.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4890
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
      Height          =   4665
      Left            =   150
      TabIndex        =   4
      Top             =   120
      Width           =   6555
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   570
         Top             =   3390
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         DefaultExt      =   "doc"
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
         Height          =   1545
         Left            =   240
         TabIndex        =   5
         Top             =   690
         Width           =   5955
         Begin VB.TextBox txtcodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   0
            Left            =   2730
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
            Left            =   2730
            MaxLength       =   1
            TabIndex        =   1
            Top             =   870
            Width           =   330
         End
         Begin VB.Image imgFec 
            Height          =   240
            Index           =   0
            Left            =   2430
            Picture         =   "frmTrasPoste.frx":000C
            ToolTipText     =   "Buscar fecha"
            Top             =   510
            Width           =   240
         End
         Begin VB.Label Label4 
            Caption         =   "Fecha"
            ForeColor       =   &H00972E0B&
            Height          =   255
            Index           =   16
            Left            =   1500
            TabIndex        =   7
            Top             =   540
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
            Left            =   1500
            TabIndex        =   6
            Top             =   900
            Width           =   645
         End
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   4905
         TabIndex        =   3
         Top             =   3780
         Width           =   975
      End
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   3720
         TabIndex        =   2
         Top             =   3780
         Width           =   975
      End
      Begin MSComctlLib.ProgressBar Pb1 
         Height          =   255
         Left            =   210
         TabIndex        =   8
         Top             =   2730
         Width           =   6030
         _ExtentX        =   10636
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Label lblProgres 
         Height          =   285
         Index           =   0
         Left            =   180
         TabIndex        =   10
         Top             =   3120
         Width           =   6195
      End
      Begin VB.Label lblProgres 
         Height          =   285
         Index           =   1
         Left            =   180
         TabIndex        =   9
         Top             =   3480
         Width           =   6195
      End
   End
End
Attribute VB_Name = "frmTrasPoste"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' PROGRAMA DE TRASPASO DE POSTE PARA ALZICOOP
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
Dim cad As String
Dim cadTABLA As String

Dim vContad As Long

Dim PrimeraVez As Boolean

Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, 0, cerrar
    If cerrar Then Unload Me
End Sub

Private Sub cmdAceptar_Click()
Dim sql As String
Dim i As Byte
Dim cadwhere As String
Dim b As Boolean
Dim nomfic As String
Dim cadena As String
Dim cadena1 As String

On Error GoTo eError
    If Not DatosOk Then Exit Sub
    
    Me.CommonDialog1.InitDir = App.path & "\TPV\V"
    Me.CommonDialog1.DefaultExt = Format(txtcodigo(1).Text, "000")
'    Me.CommonDialog1.Filter = Format(txtcodigo(1).Text, "000")
    cadena = Format(CDate(txtcodigo(0).Text), FormatoFecha)
    CommonDialog1.Filter = "Archivos BV|BV" & Mid(cadena, 9, 2) & Mid(cadena, 6, 2) & Mid(cadena, 3, 2) & "." & Format(txtcodigo(1).Text, "000") & "|"
    CommonDialog1.FilterIndex = 1

    
'    Me.CommonDialog1.FileName = "BV" & Format(CDate(txtcodigo(0).Text), "ddmmyy") & "." & Format(txtcodigo(1).Text, "000")
    Me.CommonDialog1.ShowOpen
    
    If Me.CommonDialog1.FileName <> "" Then
        InicializarVbles
        InicializarTabla
            '========= PARAMETROS  =============================
        'Añadir el parametro de Empresa
        cadParam = cadParam & "|pEmpresa=""" & vEmpresa.nomEmpre & """|"
        numParam = numParam + 1

        If FicheroCorrecto(0) Then
            
          If Dir(Replace(Me.CommonDialog1.FileName, "BV", "BO")) = "" Then
            If MsgBox("No se ha encontrado el fichero BO correspondiente. ¿Desea continuar? ", vbQuestion + vbYesNo + vbDefaultButton1) = vbNo Then Exit Sub
          End If
          
          If ProcesarFichero2(Me.CommonDialog1.FileName) Then
                cadTABLA = "tmpinformes"
                cadFormula = "{tmpinformes.codusu} = " & vSesion.Codigo
                
                sql = "select count(*) from tmpinformes where codusu = " & vSesion.Codigo
                
                If TotalRegistros(sql) <> 0 Then
'                If HayRegParaInforme(cadTABLA, cadSelect) Then
                    MsgBox "Hay errores en el Traspaso de Postes. Debe corregirlos previamente.", vbExclamation
                    cadTitulo = "Errores de Traspaso de Poste"
                    cadNombreRPT = "rErroresTrasPoste.rpt"
                    LlamarImprimir
                    Exit Sub
                Else
                    Conn.BeginTrans
                    b = ProcesarFichero(Me.CommonDialog1.FileName)
                    If FicheroCorrecto(1) And b Then
'
'  BV y BO se dejaran en el mismo directorio
'                        nomfic = Replace(Me.CommonDialog1.FileName, "\V\", "\T\")
'                        nomfic = Replace(Me.CommonDialog1.FileName, "\v\", "\t\")
                        nomfic = Me.CommonDialog1.FileName
                        If Dir(Replace(nomfic, "BV", "BO")) <> "" Then
                            b = ProcesarFichero(Replace(nomfic, "BV", "BO"))
                        End If
                    End If
                End If
          End If
        Else
            MsgBox "El fichero no se corresponde con la Fecha y Turno introducidas. Revise.", vbExclamation
            Exit Sub
        End If
    Else
        MsgBox "No ha seleccionado ningún fichero", vbExclamation
        Exit Sub
    End If
             
             
eError:
    If Err.Number <> 0 Or Not b Then
        Conn.RollbackTrans
        MsgBox "No se ha podido realizar el proceso. LLame a Ariadna.", vbExclamation
    Else
        Conn.CommitTrans
        MsgBox "Proceso realizado correctamente.", vbExclamation
        Pb1.visible = False
        lblProgres(0).Caption = ""
        lblProgres(1).Caption = ""
        cmdCancel_Click
    End If
    
'    cadTABLA = "tmpinformes"
'    cadFormula = "{tmpinformes.codusu} = " & vSesion.Codigo
'
'    If HayRegParaInforme(cadTABLA, cadSelect) Then
'          cadTitulo = "Errores de Traspaso de Poste"
'          cadNombreRPT = "rErroresTrasPoste.rpt"
'          LlamarImprimir
'    End If
End Sub

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
     txtcodigo(0).Text = Format(Now - 1, "dd/mm/yyyy")

    '###Descomentar
'    CommitConexion
         
    FrameCobrosVisible True, h, w
    Pb1.visible = False
    
    'Esto se consigue poneinedo el cancel en el opcion k corresponda
    Me.cmdCancel.Cancel = True
'   Me.Width = w + 70
'   Me.Height = h + 350
End Sub

Private Sub Form_Unload(Cancel As Integer)
    DesBloqueoManual ("TRASPOST")
End Sub

Private Sub frmC_Selec(vFecha As Date)
 'Fecha
    txtcodigo(CByte(imgFec(0).Tag)).Text = Format(vFecha, "dd/MM/yyyy")
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
    imgFec(0).Tag = Index 'independentment de les dates que tinga, sempre pose l'index en la 27
    If txtcodigo(Index).Text <> "" Then frmC.NovaData = txtcodigo(Index).Text

    frmC.Show vbModal
    Set frmC = Nothing
    PonerFoco txtcodigo(CByte(imgFec(0).Tag) + 1)
    ' ***************************
End Sub

Private Sub ListView1_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
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
'15/02/2007
'    KEYpress KeyAscii
'ahora
    If KeyAscii = teclaBuscar Then
        Select Case Index
            Case 0: KEYFecha KeyAscii, 0 'fecha
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
Dim cad As String, cadTipo As String 'tipo cliente

    'Quitar espacios en blanco por los lados
    txtcodigo(Index).Text = Trim(txtcodigo(Index).Text)
    
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub

    
    Select Case Index
        Case 0 'FECHAS
            If txtcodigo(Index).Text <> "" Then PonerFormatoFecha txtcodigo(Index)
            
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

 

Private Function DatosOk() As Boolean
Dim b As Boolean
Dim sql As String
   b = True

   If txtcodigo(0).Text = "" And b Then
        MsgBox "El campo fecha debe de tener un valor. Reintroduzca.", vbExclamation
        b = False
        PonerFoco txtcodigo(0)
    End If
    
    If txtcodigo(1).Text = "" And b Then
        MsgBox "El número de Turno debe de tener un valor. Reintroduzca.", vbExclamation
        b = False
        PonerFoco txtcodigo(1)
    End If
 
    ' COMPROBAMOS QUE EL TRASPASO DE POSTES NO HAYA SIDO HECHO ANTERIORMENTE
    If b Then
        sql = "SELECT count(*) FROM sturno WHERE fechatur = " & DBSet(txtcodigo(0).Text, "F") & _
              " AND codturno = " & DBSet(txtcodigo(1).Text, "N")
        If TotalRegistros(sql) <> 0 Then
            MsgBox "Este Turno ya ha sido traspasado. Reintroduzca.", vbExclamation
            b = False
            PonerFoco txtcodigo(1)
        End If
    End If
 
    DatosOk = b
End Function



Private Function RecuperaFichero() As Boolean
Dim nf As Integer

    RecuperaFichero = False
    nf = FreeFile
    Open App.path For Input As #nf ' & "\BV" & Format(CDate(txtcodigo(0).Text), "ddmmyy") & "." & Format(txtcodigo(1).Text, "000") For Input As #NF
    Line Input #nf, cad
    Close #nf
    If cad <> "" Then RecuperaFichero = True
    
End Function


Private Function ProcesarFichero(nomfich As String) As Boolean
Dim nf As Long
Dim cad As String
Dim i As Integer
Dim longitud As Long
Dim rs As adodb.Recordset
Dim Rs1 As adodb.Recordset
Dim NumReg As Long
Dim sql As String
Dim SQL1 As String
Dim Total As Long
Dim v_cant As Currency
Dim v_impo As Currency
Dim v_prec As Currency
Dim b As Boolean

    ProcesarFichero = False
    nf = FreeFile
    
    Open nomfich For Input As #nf ' & "\BV" & Format(CDate(txtcodigo(0).Text), "ddmmyy") & "." & Format(txtcodigo(1).Text, "000") For Input As #NF
    
    Line Input #nf, cad
    i = 0
    
    lblProgres(0).Caption = "Procesando Fichero: " & nomfich
    longitud = FileLen(nomfich)
    
    Pb1.visible = True
    Me.Pb1.Max = longitud
    Me.Refresh
    Me.Pb1.Value = 0
    ' PROCESO DEL FICHERO BV
    If Mid(NombreFichero(nomfich), 1, 2) = "BV" Then
        While Not EOF(nf)
            i = i + 1
            Me.Pb1.Value = Me.Pb1.Value + Len(cad)
            lblProgres(1).Caption = "Linea " & i
            Me.Refresh
            b = True
            Select Case Mid(cad, 1, 1)
                Case "C" ' CABECERA
                    b = InsertarCabecera(cad)
                Case "L" ' LINEAS
                    b = InsertarLinea(cad)
                Case "R" ' RECAUDACION
                    b = InsertarRecaudacion(cad)
                Case "X" ' SALIDAS
                    b = InsertarSalida(cad)
                Case Else
                
            End Select
            
            If b = False Then
                ProcesarFichero = False
                Exit Function
            End If
            
            Line Input #nf, cad
        Wend
        Close #nf
    Else
    ' PROCESO DEL FICHERO BO
        While Not EOF(nf)
            i = i + 1
            
               
            Me.Pb1.Value = Me.Pb1.Value + Len(cad)
            lblProgres(1).Caption = "Linea " & i
            Me.Refresh

            InsertarLineaTurno (cad)
            
            Line Input #nf, cad
        Wend
        If i > 1 Then InsertarLineaTurno (cad)
        
    
        Close #nf
        
        sql = "select count(*) from scaalb, sfamia, sartic where scaalb.fecalbar = " & DBSet(txtcodigo(0).Text, "F") & _
              " and codturno = " & DBSet(txtcodigo(1).Text, "N") & " and scaalb.codartic = sartic.codartic and " & _
              " sfamia.tipfamia = 0 and " & _
              " sartic.codfamia = sfamia.codfamia"
              
        Total = TotalRegistros(sql)
        i = 0
        
        sql = "select codartic, sum(litrosve), sum(importel) from sturno where fechatur = " & DBSet(txtcodigo(0).Text, "F") & _
              " and codturno = " & DBSet(txtcodigo(1).Text, "N") & " and tiporegi = 2 group by codartic order by codartic "
              
        Set rs = New adodb.Recordset
        rs.Open sql, Conn, adOpenDynamic, adLockOptimistic
        If Not rs.EOF Then rs.MoveFirst
        
        While Not rs.EOF
            Set Rs1 = New adodb.Recordset
            SQL1 = "select sum(cantidad), sum(importel) from scaalb where fecalbar = " & DBSet(txtcodigo(0).Text, "F") & _
                   " and codturno = " & DBSet(txtcodigo(1).Text, "N") & " and codartic = " & DBLet(rs!codArtic)
                   
            Rs1.Open SQL1, Conn, adOpenDynamic, adLockOptimistic
            If Not Rs1.EOF Then Rs1.MoveFirst
            
            If Total <> 0 Then
                If DBLet(rs.Fields(2).Value) > DBLet(Rs1.Fields(1).Value) Then
                    i = i + 1
                    v_cant = DBLet(rs.Fields(1).Value) - DBLet(Rs1.Fields(0).Value)
                    v_impo = DBLet(rs.Fields(2).Value) - DBLet(Rs1.Fields(1).Value)
                    v_prec = 0
                    If v_cant <> 0 Then v_prec = Round2(v_impo / v_cant, 3)
                
                    ' insertamos en la tabla de albaranes
                    NumReg = SugerirCodigoSiguienteStr("scaalb", "codclave")
                    sql = "INSERT INTO scaalb (codclave, codsocio, numalbar, fecalbar, horalbar, " & _
                          "codturno, codartic, cantidad, preciove , importel, codforpa, " & _
                          "NumFactu , NumLinea) VALUES (" & DBSet(NumReg, "N") & ", 0, 'MANUAL'," & DBSet(txtcodigo(0).Text, "F") & "," & _
                          DBSet(txtcodigo(0).Text & " " & Time, "FH") & "," & DBSet(txtcodigo(1).Text, "N") & "," & DBLet(rs.Fields(0).Value) & "," & _
                          DBSet(v_cant, "N") & "," & DBSet(v_prec, "N") & "," & DBSet(v_impo, "N") & ",1,0," & _
                          DBSet(i, "N") & ")"
                    Conn.Execute sql
                End If
            End If
            
            Set Rs1 = Nothing
            rs.MoveNext
        Wend
        
        Set rs = Nothing
    End If
    
    If cad <> "" Then ProcesarFichero = True
    
    Pb1.visible = False
    lblProgres(0).Caption = ""
    lblProgres(1).Caption = ""

End Function
                
Private Function ProcesarFichero2(nomfich As String) As Boolean
Dim nf As Long
Dim cad As String
Dim i As Integer
Dim longitud As Long
Dim rs As adodb.Recordset
Dim Rs1 As adodb.Recordset
Dim NumReg As Long
Dim sql As String
Dim SQL1 As String
Dim Total As Long
Dim v_cant As Currency
Dim v_impo As Currency
Dim v_prec As Currency

    On Error GoTo eProcesarFichero2
    
    ProcesarFichero2 = False
    nf = FreeFile
    Open nomfich For Input As #nf ' & "\BV" & Format(CDate(txtcodigo(0).Text), "ddmmyy") & "." & Format(txtcodigo(1).Text, "000") For Input As #NF
    
    Line Input #nf, cad
    i = 0
    
    lblProgres(0).Caption = "Comprobando Tarjetas Socios: " & nomfich
    longitud = FileLen(nomfich)
    
    Pb1.visible = True
    Me.Pb1.Max = longitud
    Me.Refresh
    Me.Pb1.Value = 0
    ' PROCESO DEL FICHERO BV
    If Mid(NombreFichero(nomfich), 1, 2) = "BV" Then
'        If i = 235 Then Stop

        While Not EOF(nf)
            i = i + 1
            
            Me.Pb1.Value = Me.Pb1.Value + Len(cad)
            lblProgres(1).Caption = "Linea " & i
            Me.Refresh
            Select Case Mid(cad, 1, 1)
                Case "C" ' CABECERA
                    InsertarCabecera2 (cad)
                Case "L" ' LINEAS
                Case "R" ' RECAUDACION
                Case "X" ' SALIDAS
                Case Else
                
            End Select
            Line Input #nf, cad
        Wend
        Close #nf
    End If
    
    If cad <> "" Then ProcesarFichero2 = True
    
    Pb1.visible = False
    lblProgres(0).Caption = ""
    lblProgres(1).Caption = ""

eProcesarFichero2:
    If Err.Number <> 0 Then
        MsgBox "Se ha producido un error en el proceso de comprobación", vbExclamation
    End If
End Function
                
                
                
                
                
Private Function InsertarCabecera(cad As String) As Boolean
Dim numfactu As String
Dim TipDocu As String
Dim FechaCa As String
Dim turno As String
Dim hora As String
Dim forpa As String
Dim Tarje As String
Dim Tarje1 As String
Dim Matric As String
Dim NomCli As String
Dim NifCli As String
Dim Ticket As String
Dim CtaConta As String ' cuenta contable de clientes contado
Dim codsoc As String
Dim sql As String

    On Error GoTo eInsertarCabecera

    InsertarCabecera = False

    numfactu = 0
    TipDocu = Mid(cad, 10, 1)
    FechaCa = Mid(cad, 11, 2) & Mid(cad, 13, 2) & "20" & Mid(cad, 15, 2)
    turno = Mid(cad, 17, 1)
    hora = Mid(cad, 18, 2) & ":" & Mid(cad, 21, 2) & ":00"
    forpa = Mid(cad, 49, 2)
    Tarje = Mid(cad, 53, 7)
    Tarje1 = Mid(cad, 60, 5)
    Matric = Mid(cad, 65, 10)
    NomCli = Mid(cad, 91, 25)
    NifCli = Mid(cad, 116, 9)
            
    '06/03/2007 añadida estas 2 lineas que faltaba
    If CInt(forpa) <> 2 And Trim(Tarje) <> Trim(Tarje1) Then Tarje = Tarje1
    If Tarje = "" Then Tarje = "0"
    
    Select Case TipDocu
        Case "O"
            Ticket = Mid(cad, 2, 8)
        Case "T"
            Ticket = Mid(cad, 23, 8)
        Case "A"
            Ticket = Mid(cad, 31, 8)
        Case "F"
            Ticket = Mid(cad, 2, 8)
            numfactu = Mid(cad, 39, 8)
        
            'SOLAMENTE EN EL CASO DE QUE SEA FACTURA COMPRUEBO QUE EXISTA EL NIF DEL SOCIO
            'Y SI NO EXISTE INTRODUCIRLO EN LA TABLA DE SOCIOS Y TARJETAS
            Tarje = DevuelveDesdeBDNew(cPTours, "ssocio", "codsocio", "nifsocio", NifCli, "T")
            If Tarje = "" Then
                Tarje = 900000
                Tarje = SugerirCodigoSiguienteStr("ssocio", "codsocio", "codsocio >= 900000 and codsocio <= 999998")
                
                CtaConta = ""
                CtaConta = DevuelveDesdeBD("ctaconta", "sparam", "codparam", "0", "N")
                
                sql = "INSERT INTO ssocio (codsocio, codcoope, nomsocio, domsocio, codposta, pobsocio, " & _
                      "prosocio, nifsocio, telsocio, faxsocio, movsocio, maisocio, wwwsocio, fechaalt, " & _
                      "fechabaj, codtarif, codbanco, codsucur, digcontr, cuentaba, impfactu, dtolitro, " & _
                      "codforpa, tipsocio, bonifbas, bonifesp, codsitua, codmacta, obssocio) VALUES (" & _
                      DBSet(Tarje, "N") & ",0," & DBSet(NomCli, "T") & ",'DESCONOCIDA','46','VALENCIA', " & _
                      "'VALENCIA'," & DBSet(NifCli, "T") & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & _
                      DBSet(txtcodigo(0).Text, "F") & "," & _
                      ValorNulo & ",0," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & ",0,0," & _
                      "0,0,0,0,0," & ValorNulo & "," & ValorNulo & ")"
                      
                Conn.Execute sql
                      
                sql = "INSERT INTO starje (codsocio, numlinea, numtarje, nomtarje, codbanco, codsucur, " & _
                      "digcontr, cuentaba, tiptarje) VALUES (" & DBSet(Tarje, "N") & ",1," & DBSet(NomCli, "T") & "," & ValorNulo & "," & ValorNulo & "," & _
                      ValorNulo & "," & ValorNulo & ",0)"
                
                Conn.Execute sql
            End If
    End Select
   

    'MIRAMOS SI EXISTE LA TARJETA
    codsoc = ""
    codsoc = DevuelveDesdeBD("codsocio", "starje", "numtarje", Tarje, "T")
    If Tarje = "       " Then Tarje = "0000000"
    If codsoc = "" Then
    
        sql = "insert into tmpinformes (codusu, importe1, fecha1, campo1, campo2, importe2, nombre1) values (" & _
              vSesion.Codigo & "," & DBSet(Ticket, "N") & ",'" & Mid(FechaCa, 5, 4) & Mid(FechaCa, 3, 2) & Mid(FechaCa, 1, 2) & "'," & DBSet(Format(hora, "hh"), "N") & _
              "," & DBSet(Format(hora, "mm"), "N") & "," & DBSet(Tarje, "N") & ",'Nro. Tarjeta no existe') "
              
        Conn.Execute sql
        
        
    Else
        sql = "update scaalb set codsocio = " & DBSet(codsoc, "N") & ", numtarje = " & DBSet(Tarje, "N") & ", numalbar = " & _
               DBSet(Ticket, "T") & ", horalbar = " & DBSet(txtcodigo(0).Text & " " & hora, "FH") & ", matricul = " & DBSet(Matric, "T") & _
               ", codforpa = " & DBSet(forpa, "N") & ", numfactu = " & DBSet(numfactu, "N") & _
               " where fecalbar = " & DBSet(txtcodigo(0).Text, "F") & " and codturno = " & DBSet(txtcodigo(1).Text, "N") & _
               " and numalbar = " & DBSet(vContad, "T")
               
        Conn.Execute sql
    End If
    
    vContad = vContad + 1

    InsertarCabecera = True
    
eInsertarCabecera:
    If Err.Number <> 0 Then
        MsgBox "Error en Insertar Cabecera " & Err.Description, vbExclamation
    End If

End Function
            
Private Sub InsertarCabecera2(cad As String)
Dim numfactu As String
Dim TipDocu As String
Dim FechaCa As String
Dim turno As String
Dim hora As String
Dim forpa As String
Dim Tarje As String
Dim Tarje1 As String
Dim Matric As String
Dim NomCli As String
Dim NifCli As String
Dim Ticket As String
Dim CtaConta As String ' cuenta contable de clientes contado
Dim codsoc As String
Dim sql As String

Dim Mens As String

    numfactu = 0
    TipDocu = Mid(cad, 10, 1)
    FechaCa = Mid(cad, 11, 2) & Mid(cad, 13, 2) & "20" & Mid(cad, 15, 2)
    turno = Mid(cad, 17, 1)
    hora = Mid(cad, 18, 2) & ":" & Mid(cad, 21, 2) & ":00"
    forpa = Mid(cad, 49, 2)
    Tarje = Mid(cad, 53, 7)
    Tarje1 = Mid(cad, 60, 5)
    Matric = Mid(cad, 65, 10)
    NomCli = Mid(cad, 91, 25)
    NifCli = Mid(cad, 116, 9)
            
    Select Case TipDocu
        Case "O"
            Ticket = Mid(cad, 2, 8)
        Case "T"
            Ticket = Mid(cad, 23, 8)
        Case "A"
            Ticket = Mid(cad, 31, 8)
        Case "F"
            Ticket = Mid(cad, 2, 8)
            numfactu = Mid(cad, 39, 8)
        
            'SOLAMENTE EN EL CASO DE QUE SEA FACTURA COMPRUEBO QUE EXISTA EL NIF DEL SOCIO
            'Y SI NO EXISTE INTRODUCIRLO EN LA TABLA DE SOCIOS Y TARJETAS
            Tarje = DevuelveDesdeBDNew(cPTours, "ssocio", "codsocio", "nifsocio", NifCli, "T")
            If Tarje = "" Then
                Tarje = 900000
                Tarje = SugerirCodigoSiguienteStr("ssocio", "codsocio", "codsocio >= 900000 and codsocio <= 999998")
                
'                CtaConta = ""
'                CtaConta = DevuelveDesdeBD("ctaconta", "sparam", "codparam", "01", "N")
                
                sql = "INSERT INTO ssocio (codsocio, codcoope, nomsocio, domsocio, codposta, pobsocio, " & _
                      "prosocio, nifsocio, telsocio, faxsocio, movsocio, maisocio, wwwsocio, fechaalt, " & _
                      "fechabaj, codtarif, codbanco, codsucur, digcontr, cuentaba, impfactu, dtolitro, " & _
                      "codforpa, tipsocio, bonifbas, bonifesp, codsitua, codmacta, obssocio) VALUES (" & _
                      DBSet(Tarje, "N") & ",0," & DBSet(NomCli, "T") & ",'DESCONOCIDA','46','VALENCIA', " & _
                      "'VALENCIA'," & DBSet(NifCli, "T") & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & _
                      DBSet(txtcodigo(0).Text, "F") & "," & _
                      ValorNulo & ",0," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & ",0,0," & _
                      "0,0,0,0,0," & DBSet(vParamAplic.CtaContable, "T") & "," & ValorNulo & ")"
                      
                Conn.Execute sql
                      
                sql = "INSERT INTO starje (codsocio, numlinea, numtarje, nomtarje, codbanco, codsucur, " & _
                      "digcontr, cuentaba, tiptarje) VALUES (" & DBSet(Tarje, "N") & ",1," & DBSet(Tarje, "N") & "," & DBSet(NomCli, "T") & "," & ValorNulo & "," & ValorNulo & "," & _
                      ValorNulo & "," & ValorNulo & ",0)"
                
                Conn.Execute sql
            End If
    End Select
   

    'MIRAMOS SI EXISTE LA TARJETA
    codsoc = ""
    codsoc = DevuelveDesdeBD("codsocio", "starje", "numtarje", Tarje, "T")
    If Tarje = "       " Then Tarje = "0000000"
    If codsoc = "" Then
        Mens = "Nro. Tarjeta no existe"
        sql = "insert into tmpinformes (codusu, importe1, fecha1, campo1, campo2, importe2, nombre1) values (" & _
              vSesion.Codigo & "," & DBSet(Ticket, "N") & ",'" & Mid(FechaCa, 5, 4) & Mid(FechaCa, 3, 2) & Mid(FechaCa, 1, 2) & "'," & DBSet(Format(hora, "hh"), "N") & _
              "," & DBSet(Format(hora, "mm"), "N") & "," & DBSet(Tarje, "N") & "," & DBSet(Mens, "T") & ")"
              
        Conn.Execute sql
    End If

End Sub
            
            
            
            
            
Private Function InsertarLinea(cad As String) As Boolean
Dim NumLin As String
Dim codpro As String
Dim articulo As String
Dim Familia As String
Dim cantidad As String
Dim precio As String
Dim Importe As String
Dim ImpDes As String
Dim CodIVA As String
Dim b As Boolean
Dim Codclave As String
Dim sql As String

Dim Import As Currency

    On Error GoTo eInsertarLinea

    InsertarLinea = False
    
    NumLin = Mid(cad, 2, 2)
    codpro = Mid(cad, 4, 15)
    
    If Not EsNumerico2(codpro) Then
        codpro = Mid(cad, 17, 2)
    End If
    articulo = Mid(cad, 19, 35)
    Familia = Mid(cad, 44, 2)
    
    cantidad = Mid(cad, 55, 6) & "," & Mid(cad, 61, 2)
    precio = Mid(cad, 63, 5) & "," & Mid(cad, 68, 3)
    ImpDes = Mid(cad, 79, 5) & "," & Mid(cad, 85, 2)
    Importe = Mid(cad, 87, 6) & "," & Mid(cad, 93, 2)
    CodIVA = Mid(cad, 95, 1)
    
    '### [Monica] 17/09/2007
    'no insertamos aquellas lineas de albaran de importe = 0
    Import = DBSet(Importe, "N")
    If Import = 0 Then
        InsertarLinea = True
        Exit Function
    End If
    'hasta aqui
    
    b = InsertarFamiliaSiNoExiste(Familia)

    If b Then
         b = InsertarArticuloSiNoExiste(codpro, Familia, articulo, precio, CodIVA)
         If b Then
            'VRS:4.0.1(0) actualizamos el precio de articulo
            sql = "update sartic set preventa = " & DBSet(precio, "N") & _
                  " where codartic = " & DBSet(codpro, "N")
            Conn.Execute sql
         
            ' insertamos en la tabla de albaranes
            Codclave = SugerirCodigoSiguienteStr("scaalb", "codclave")

            sql = "INSERT INTO scaalb (codclave, codsocio, numalbar, fecalbar, horalbar, " & _
                   "codturno, codartic, cantidad, preciove, importel, codforpa, " & _
                   "numfactu, numlinea) VALUES (" & DBSet(Codclave, "T") & ",0," & _
                   DBSet(vContad, "T") & "," & DBSet(txtcodigo(0).Text, "F") & "," & DBSet(txtcodigo(0).Text & " " & Time, "FH") & "," & DBSet(txtcodigo(1).Text, "N") & "," & _
                   DBSet(codpro, "N") & "," & DBSet(cantidad, "N") & "," & DBSet(precio, "N") & "," & _
                   DBSet(Importe, "N") & ", 0, 0," & DBSet(NumLin, "N") & ")"

            Conn.Execute sql
         End If
    End If
    InsertarLinea = True
eInsertarLinea:
    If Err.Number <> 0 Then
        MsgBox "Error en Insertar Linea " & Err.Description, vbExclamation
    End If
End Function
            
Private Function InsertarRecaudacion(cad As String) As Boolean
Dim forpa As String
Dim Importe As String
Dim sql As String

    On Error GoTo eInsertarRecaudacion

    InsertarRecaudacion = False
    forpa = Mid(cad, 2, 2)
    Importe = Mid(cad, 14, 8) & "," & Mid(cad, 22, 2)
    If CCur(Importe) <> 0 Then
        sql = "insert into srecau (fechatur, codturno, codforpa, importel, intconta) values (" & _
              DBSet(txtcodigo(0).Text, "F") & "," & DBSet(txtcodigo(1).Text, "N") & "," & _
              DBSet(CInt(forpa), "N") & "," & DBSet(Importe, "N") & ",0)"
    
        Conn.Execute sql
    End If
    InsertarRecaudacion = True
eInsertarRecaudacion:
    If Err.Number <> 0 Then
        MsgBox "Error en Insertar Recaudacion en " & Err.Description, vbExclamation
    End If
    
End Function

Private Function InsertarSalida(cad As String) As Boolean
Dim TipMov As String
Dim Importe As Currency
Dim sql As String
Dim i  As Integer

    On Error GoTo eInsertarSalida
    
    
    InsertarSalida = False
    TipMov = Mid(cad, 2, 6)
    i = InStr(Mid(cad, 8, 10), "-")
    If i = 0 Then
        Importe = Format(CCur(TransformaPuntosComas(Mid(cad, 8, 10))), "######0.00")
    Else
        Importe = Format(CCur(Replace(TransformaPuntosComas(Mid(cad, 8, 10)), "-", "") * (-1)), "######0.00")
    End If
    
    If TipMov = "MOVIMI" And CCur(Importe) <> 0 Then
        sql = "insert into srecau (fechatur, codturno, codforpa, importel, intconta) values (" & _
              DBSet(txtcodigo(0).Text, "F") & "," & DBSet(txtcodigo(1).Text, "N") & "," & _
              "99, " & DBSet(Importe, "N") & ",0)"
              
        Conn.Execute sql
    End If
    InsertarSalida = True
eInsertarSalida:
    If Err.Number <> 0 Then
        MsgBox "Error en Insertar Salida en " & Err.Description, vbExclamation
    End If
End Function

Private Sub InsertarLineaTurno(cad As String)
Dim codpro As String
Dim cantidad As String
Dim precio As String
Dim Importe As String
Dim sql As String
Dim NumLin As Long
Dim cWhere As String


    codpro = Mid(cad, 35, 2)
    cantidad = Mid(cad, 54, 6) & "," & Mid(cad, 60, 2)
    precio = Mid(cad, 42, 2) & "," & Mid(cad, 44, 2)
    Importe = Mid(cad, 47, 5) & "," & Mid(cad, 52, 2)
    
    sql = ""
    sql = DevuelveDesdeBDNew(cPTours, "sturno", "codturno", "fechatur", txtcodigo(0).Text, "F", , "codturno", txtcodigo(1).Text, "N", "codartic", codpro, "N")
    If sql = "" Then
    
        cWhere = "fechatur=" & DBSet(txtcodigo(0).Text, "F") & " and codturno = " & DBSet(txtcodigo(1).Text, "N")
        NumLin = CLng(SugerirCodigoSiguienteStr("sturno", "numlinea", cWhere))
        'insertamos
        sql = "INSERT INTO sturno (fechatur, codturno, numlinea, tiporegi, numtanqu, nummangu, " & _
              " codartic, litrosve, importel, containi, contafin, tipocred) VALUES (" & _
              DBSet(txtcodigo(0).Text, "F") & "," & DBSet(txtcodigo(1).Text, "N") & "," & DBSet(NumLin, "N") & ",2,1,1," & _
              DBSet(codpro, "N") & "," & DBSet(cantidad, "N") & "," & DBSet(Importe, "N") & ",0,0,0)"
              
        Conn.Execute sql
    Else
        'actualizamos
        sql = "UPDATE sturno SET importel = importel + " & DBSet(Importe, "N") & ", litrosve = litrosve +  " & DBSet(cantidad, "N") & " WHERE fechatur = " & _
              DBSet(txtcodigo(0).Text, "F") & " AND codturno = " & DBSet(txtcodigo(1).Text, "N") & " AND codartic = " & _
              DBSet(codpro, "N")
              
        Conn.Execute sql
    End If
End Sub

Private Function FicheroCorrecto(tipo As String) As Boolean
Dim fic As String
    
    FicheroCorrecto = False
    
    If tipo = "0" Then
        fic = NombreFichero(Me.CommonDialog1.FileName) = ("BV" & Format(CDate(txtcodigo(0).Text), "ddmmyy") & "." & Format(txtcodigo(1).Text, "000"))
    Else
        fic = NombreFichero(Replace(Me.CommonDialog1.FileName, "BV", "BO")) = ("BO" & Format(CDate(txtcodigo(0).Text), "ddmmyy") & "." & Format(txtcodigo(1).Text, "000"))
' BV y BO estan en el mismo directorio
'        fic = NombreFichero(Replace(Replace(Me.CommonDialog1.FileName, "BV", "BO"), "\v\", "\t\")) = ("BO" & Format(CDate(txtcodigo(0).Text), "ddmmyy") & "." & Format(txtcodigo(1).Text, "000"))
    End If
    
    FicheroCorrecto = fic
    
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

Private Sub InicializarTabla()
Dim sql As String
    sql = "delete from tmpinformes where codusu = " & vSesion.Codigo
    
    Conn.Execute sql
End Sub
