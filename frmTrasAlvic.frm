VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmTrasAlvic 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Traspaso Datos Poste"
   ClientHeight    =   4890
   ClientLeft      =   45
   ClientTop       =   2430
   ClientWidth     =   6825
   Icon            =   "frmTrasAlvic.frx":0000
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
      Begin VB.CheckBox Check1 
         Caption         =   "Procesar el fichero de Compras"
         Height          =   195
         Left            =   270
         TabIndex        =   11
         Top             =   2400
         Width           =   2535
      End
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
            Top             =   495
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
            Picture         =   "frmTrasAlvic.frx":000C
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
         Left            =   3690
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
      Begin VB.Image imgAyuda 
         Height          =   240
         Index           =   0
         Left            =   2910
         MousePointer    =   4  'Icon
         Tag             =   "-1"
         ToolTipText     =   "Ayuda"
         Top             =   2370
         Width           =   240
      End
      Begin VB.Label lblProgres 
         Height          =   375
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
Attribute VB_Name = "frmTrasAlvic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' PROGRAMA DE TRASPASO DE POSTE (Alvic) PARA ALZICOOP y Regaixo
' basado en frmTrasPoste ( de Alzira )
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
Dim Sql As String
Dim i As Byte
Dim cadWHERE As String
Dim b As Boolean
Dim NomFic As String
Dim Cadena As String
Dim cadena1 As String

On Error GoTo eError


    If Not DatosOk Then Exit Sub
    
    If vParamAplic.Cooperativa = 2 Then
        TraspasoRegaixo
        Unload Me
        Exit Sub
    End If
    
    Me.CommonDialog1.DefaultExt = "TXT"
    Cadena = Format(CDate(txtCodigo(0).Text), FormatoFecha)
    CommonDialog1.FilterIndex = 1
    CommonDialog1.CancelError = True
    If vParamAplic.Cooperativa = 5 Then
        Me.CommonDialog1.FileName = "ventas" & ".txt"
    Else
        Me.CommonDialog1.FileName = "ventas" & Mid(txtCodigo(0), 1, 2) & Mid(txtCodigo(0), 4, 2) & Mid(txtCodigo(0), 9, 2) & ".txt"
    End If
    Me.CommonDialog1.ShowOpen
    
    If Me.CommonDialog1.FileName <> "" Then
        InicializarVbles
        InicializarTabla
            '========= PARAMETROS  =============================
        'Añadir el parametro de Empresa
        cadParam = cadParam & "|pEmpresa=""" & vEmpresa.nomEmpre & """|"
        numParam = numParam + 1

'[Monica]10/11/2010 añadimos las compras en alzira
        If Dir(Replace(Me.CommonDialog1.FileName, "ventas", "compras"), vbArchive) <> "" And Check1.Value Then
            If Not ComprobarFechaAlbaran(Replace(Me.CommonDialog1.FileName, "ventas", "compras")) Then
                Sql = "select count(*) from tmpinformes where codusu = " & vSesion.Codigo
                
                If TotalRegistros(Sql) <> 0 Then
                    If MsgBox("Hay albaranes de compra con fecha distinta a la del turno. ¿ Desea continuar con ventas ? ", vbQuestion + vbYesNo + vbDefaultButton1) = vbNo Then
                        Exit Sub
                    End If
                End If
            Else
                If Not ProcesarFicheroCompras(Replace(Me.CommonDialog1.FileName, "ventas", "compras")) Then
                    If MsgBox("No se ha realizado el proceso de compras. " & vbCrLf & vbCrLf & "¿ Desea continuar con el proceso de ventas ?", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
                        Pb1.visible = False
                        lblProgres(0).Caption = ""
                        lblProgres(1).Caption = ""
                        Exit Sub
                    End If
                Else
                    cadTABLA = "tmpinformes"
                    cadFormula = "{tmpinformes.codusu} = " & vSesion.Codigo
                    
                    Sql = "select count(*) from tmpinformes where codusu = " & vSesion.Codigo
                    If TotalRegistros(Sql) <> 0 Then
                        cadTitulo = "Informe de Acciones a revisar"
                        cadNombreRPT = "rInfARevisar.rpt"
                        LlamarImprimir
                    End If
                End If
            End If
        End If
        InicializarTabla
'fin
          If ProcesarFichero2(Me.CommonDialog1.FileName) Then
                cadTABLA = "tmpinformes"
                cadFormula = "{tmpinformes.codusu} = " & vSesion.Codigo
                
                Sql = "select count(*) from tmpinformes where codusu = " & vSesion.Codigo
                Sql = Sql & " and importeb1 is null "
                
                If TotalRegistros(Sql) <> 0 Then
'                If HayRegParaInforme(cadTABLA, cadSelect) Then
                    MsgBox "Hay errores en el Traspaso de Postes. Debe corregirlos previamente.", vbExclamation
                    cadTitulo = "Errores de Traspaso de Poste"
                    cadNombreRPT = "rErroresTrasPoste3.rpt"
                    LlamarImprimir
                    Exit Sub
                Else
                    Sql = "select count(*) from tmpinformes where codusu = " & vSesion.Codigo
                    Sql = Sql & " and importeb1 = 0 "
                    
                    If TotalRegistros(Sql) <> 0 Then
                        MsgBox "Hay errores en el Traspaso de Postes. Revise.", vbExclamation
                        cadTitulo = "Errores de Traspaso de Poste"
                        cadNombreRPT = "rErroresTrasPoste3.rpt"
                        LlamarImprimir
                    End If
                    
                    Conn.BeginTrans
                    b = ProcesarFichero(Me.CommonDialog1.FileName)
'                    If FicheroCorrecto(1) And b Then
''
''  BV y BO se dejaran en el mismo directorio
''                        nomfic = Replace(Me.CommonDialog1.FileName, "\V\", "\T\")
''                        nomfic = Replace(Me.CommonDialog1.FileName, "\v\", "\t\")
'                        nomfic = Me.CommonDialog1.FileName
'                        If Dir(Replace(nomfic, "BV", "BO")) <> "" Then
'                            b = ProcesarFichero(Replace(nomfic, "BV", "BO"))
'                        End If
'                    End If
                End If
          End If
'        Else
'            MsgBox "El fichero no se corresponde con la Fecha y Turno introducidas. Revise.", vbExclamation
'            Exit Sub
'        End If
    Else
        MsgBox "No ha seleccionado ningún fichero", vbExclamation
        Exit Sub
    End If
             
             
eError:
    If Err.Number <> 0 Or Not b Then
        If Err.Number = 32755 Then Exit Sub
        Conn.RollbackTrans
        MsgBox "No se ha podido realizar el proceso. LLame a Ariadna.", vbExclamation
    Else
        Conn.CommitTrans
        MsgBox "Proceso realizado correctamente.", vbExclamation
        Pb1.visible = False
        lblProgres(0).Caption = ""
        lblProgres(1).Caption = ""
        BorrarArchivo Me.CommonDialog1.FileName
        BorrarArchivo Replace(LCase(Me.CommonDialog1.FileName), "ventas", "totaliza")
        BorrarArchivo Replace(LCase(Me.CommonDialog1.FileName), "ventas", "compras")
        '[Monica]09/01/2013: nueva cooperativa de Ribarroja
        If vParamAplic.Cooperativa = 1 Or vParamAplic.Cooperativa = 5 Then
        ' solo en el caso de alzira se graba en la srecau
            BorrarArchivo Replace(LCase(Me.CommonDialog1.FileName), "ventas", "caja")
            BorrarArchivo Replace(LCase(Me.CommonDialog1.FileName), "ventas", "totales")
        End If
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

    


Private Function TraspasoRegaixo() As Boolean
Dim Sql As String
Dim b As Boolean
Dim Cadena As String
Dim i As Integer

    On Error GoTo eTraspasoRegaixo
    
    TraspasoRegaixo = False


    Me.CommonDialog1.DefaultExt = "XLS"
    Cadena = Format(CDate(txtCodigo(0).Text), FormatoFecha)
    CommonDialog1.FilterIndex = 1
    CommonDialog1.CancelError = True
    
    Me.CommonDialog1.ShowOpen
    
    If Me.CommonDialog1.FileName <> "" Then
        InicializarVbles
        InicializarTabla

        If Dir(App.path & "\trasarigasol.z") <> "" Then Kill App.path & "\trasarigasol.z"

        Shell App.path & "\trasarigasol.exe /I|" & vSesion.CadenaConexion & "|" & vSesion.Codigo & "|" & Me.CommonDialog1.FileName & "|", vbNormalFocus

            
        i = 0
        While Dir(App.path & "\trasarigasol.z") = "" And i < 300
            Me.lblProgres(0).Caption = "Procesando Insercion "
            DoEvents
            
            espera 1
            
            i = i + 1
        Wend
        
        
        If Dir(App.path & "\trasarigasol.z") = "" Then
        
            Dim NF As Integer
            NF = FreeFile
            Open App.path & "\trasarigasol.z" For Output As #NF
            Print #NF, "0"
    '        Line Input #NF, cad
            Close #NF
                    
            Unload Me
            Exit Function
        End If
        
        
        Sql = "select count(*) from tmptraspaso where codusu = " & vSesion.Codigo
        
        If TotalRegistros(Sql) <> 0 Then
    
            InicializarTabla
              If ProcesarFicheroRegaixo2() Then
                    cadTABLA = "tmpinformes"
                    cadFormula = "{tmpinformes.codusu} = " & vSesion.Codigo
                    
                    Sql = "select count(*) from tmpinformes where codusu = " & vSesion.Codigo
'                    Sql = Sql & " and importeb1 = 0 "
                    
                    If TotalRegistros(Sql) <> 0 Then
                        MsgBox "Hay errores en el Traspaso de Postes. Revise.", vbExclamation
                        cadTitulo = "Errores de Traspaso de Poste"
                        cadNombreRPT = "rErroresTrasPoste3.rpt"
                        LlamarImprimir
                        Exit Function
                    End If
                    
                    Conn.BeginTrans
                    b = ProcesarFicheroRegaixo()
                    If b Then
                        Conn.CommitTrans
                    Else
                        Conn.RollbackTrans
                    End If
              Else
                'DAVID
                
              End If
        Else
            MsgBox "No ha seleccionado ningún fichero", vbExclamation
            Exit Function
        End If
                 
    End If
    
eTraspasoRegaixo:
    If Err.Number <> 0 Or Not b Then
        If Err.Number = 32755 Then Exit Function
        
        MsgBox "No se ha podido realizar el proceso. LLame a Ariadna.", vbExclamation
    Else
        TraspasoRegaixo = True
        MsgBox "Proceso realizado correctamente.", vbExclamation
        Pb1.visible = False
        lblProgres(0).Caption = ""
        lblProgres(1).Caption = ""
        cmdCancel_Click
    End If
    
End Function


Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub Cmdleer_Click()

End Sub

Private Sub Form_Activate()
    If PrimeraVez Then
        PrimeraVez = False
        PonerFoco txtCodigo(0)
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub ActivarAyuda(sn As Boolean)
    If sn Then
        Check1.Value = 1
    Else
        Check1.Value = 0
    End If
    Check1.visible = sn ' solo si es alzira procesa el fichero de compras
    Check1.Enabled = sn
    
    imgAyuda(0).Picture = frmPpal.ImageListB.ListImages(10).Picture
    imgAyuda(0).visible = sn
    imgAyuda(0).Enabled = sn
End Sub

Private Sub Form_Load()
Dim h As Integer, w As Integer
Dim List As Collection

    PrimeraVez = True
    limpiar Me

    'IMAGES para busqueda
     txtCodigo(0).Text = Format(Now - 1, "dd/mm/yyyy")

    '###Descomentar
'    CommitConexion
    '[Monica]24/12/2015: regaixo traspasa todo el dia
    If vParamAplic.Cooperativa = 2 Then
        Label4(2).visible = False
        Label4(2).Enabled = False
        txtCodigo(1).visible = False
        txtCodigo(1).Enabled = False
    End If

         
    FrameCobrosVisible True, h, w
    Pb1.visible = False
    '[Monica]09/01/2013: Nueva cooperativa de Ribarroja
    ActivarAyuda (vParamAplic.Cooperativa = 1 Or vParamAplic.Cooperativa = 5)
    
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
    txtCodigo(CByte(imgFec(0).Tag)).Text = Format(vFecha, "dd/MM/yyyy")
End Sub

Private Sub imgAyuda_Click(Index As Integer)
Dim vCadena As String
    Select Case Index
        Case 0
           ' "____________________________________________________________"
            vCadena = "Si se ha eliminado un turno, el check ha de estar desmarcado. " & vbCrLf & vbCrLf & _
                      "El motivo es porque si se ha traspasado el fichero de compras, " & vbCrLf & _
                      "los albaranes no se eliminan cuando se borra un turno." & vbCrLf & vbCrLf
                      
                      
    End Select
    MsgBox vCadena, vbInformation, "Descripción de Ayuda"
    
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
    If txtCodigo(Index).Text <> "" Then frmC.NovaData = txtCodigo(Index).Text

    frmC.Show vbModal
    Set frmC = Nothing
    PonerFoco txtCodigo(CByte(imgFec(0).Tag) + 1)
    ' ***************************
End Sub

Private Sub ListView1_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub



Private Sub txtCodigo_GotFocus(Index As Integer)
    ConseguirFoco txtCodigo(Index), 3
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
    txtCodigo(Index).Text = Trim(txtCodigo(Index).Text)
    
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub

    
    Select Case Index
        Case 0 'FECHAS
            If txtCodigo(Index).Text <> "" Then PonerFormatoFecha txtCodigo(Index)
            
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
Dim Sql As String
   b = True

   If txtCodigo(0).Text = "" And b Then
        MsgBox "El campo fecha debe de tener un valor. Reintroduzca.", vbExclamation
        b = False
        PonerFoco txtCodigo(0)
    End If
    
    If txtCodigo(1).Text = "" And b And vParamAplic.Cooperativa <> 5 And vParamAplic.Cooperativa <> 2 Then
        MsgBox "El número de Turno debe de tener un valor. Reintroduzca.", vbExclamation
        b = False
        PonerFoco txtCodigo(1)
    End If
 
    ' COMPROBAMOS QUE EL TRASPASO DE POSTES NO HAYA SIDO HECHO ANTERIORMENTE
    If b Then
        '[Monica]10/01/2013: en la cooperativa 5 no se graba srecau
        If vParamAplic.Cooperativa = 5 Then
            Sql = "SELECT count(*) FROM scaalb WHERE fecalbar = " & DBSet(txtCodigo(0).Text, "F")
            
            If txtCodigo(1).Text <> "" Then Sql = Sql & " AND codturno = " & DBSet(txtCodigo(1).Text, "N")
            
            If TotalRegistros(Sql) <> 0 Then
                MsgBox "Este Turno ya ha sido traspasado. Reintroduzca.", vbExclamation
                b = False
                PonerFoco txtCodigo(1)
            End If
        Else
            ' faltaba comprobar que en el regaixo que no llevan turnos no se haya hecho ya el traspaso
            If vParamAplic.Cooperativa = 2 Then
                Sql = "SELECT count(*) FROM srecau WHERE fechatur = " & DBSet(txtCodigo(0).Text, "F")
                If TotalRegistros(Sql) <> 0 Then
                    MsgBox "Este Turno ya ha sido traspasado. Reintroduzca.", vbExclamation
                    b = False
                    PonerFoco txtCodigo(1)
                End If
            Else
                Sql = "SELECT count(*) FROM srecau WHERE fechatur = " & DBSet(txtCodigo(0).Text, "F") & _
                      " AND codturno = " & DBSet(txtCodigo(1).Text, "N")
                If TotalRegistros(Sql) <> 0 Then
                    MsgBox "Este Turno ya ha sido traspasado. Reintroduzca.", vbExclamation
                    b = False
                    PonerFoco txtCodigo(1)
                End If
            End If
        End If
    
    End If
 
    DatosOk = b
End Function



Private Function RecuperaFichero() As Boolean
Dim NF As Integer

    RecuperaFichero = False
    NF = FreeFile
    Open App.path For Input As #NF ' & "\BV" & Format(CDate(txtcodigo(0).Text), "ddmmyy") & "." & Format(txtcodigo(1).Text, "000") For Input As #NF
    Line Input #NF, cad
    Close #NF
    If cad <> "" Then RecuperaFichero = True
    
End Function


Private Function ProcesarFichero(nomFich As String) As Boolean
Dim NF As Long
Dim cad As String
Dim i As Integer
Dim longitud As Long
Dim RS As ADODB.Recordset
Dim Rs1 As ADODB.Recordset
Dim Numreg As Long
Dim Sql As String
Dim SQL1 As String
Dim total As Long
Dim v_cant As Currency
Dim v_impo As Currency
Dim v_prec As Currency
Dim b As Boolean
Dim NomFic As String

    ProcesarFichero = False
    NF = FreeFile
    
    Open nomFich For Input As #NF ' & "\BV" & Format(CDate(txtcodigo(0).Text), "ddmmyy") & "." & Format(txtcodigo(1).Text, "000") For Input As #NF
    
    Line Input #NF, cad
    i = 0
    
    lblProgres(0).Caption = "Procesando Fichero: " & nomFich
    longitud = FileLen(nomFich)
    
    Pb1.visible = True
    Me.Pb1.Max = longitud
    Me.Refresh
    Me.Pb1.Value = 0
        
    b = True
    While Not EOF(NF)
        i = i + 1
        
        Me.Pb1.Value = Me.Pb1.Value + Len(cad)
        lblProgres(1).Caption = "Linea " & i
        Me.Refresh
        
        '[Monica]09/01/2013: Nueva cooperativa de Ribarroja
        If vParamAplic.Cooperativa = 1 Then
            b = InsertarLineaAlz(cad)
        ElseIf vParamAplic.Cooperativa = 5 Then
            b = InsertarLineaRib(cad)
        Else
            b = InsertarLinea(cad)
        End If
        
'--monica: insertamos recaudacion leyendo de fichero al final del proceso y solo para Alzira
'        If b Then b = InsertarRecaudacion(cad)
'++monica: en regaixo hemos de insertar en srecau para la contabilizacion de turno lo habiamos quitado
'          para alzira
        
        '[Monica]09/01/2013: nueva cooperativa de Ribarroja
        If vParamAplic.Cooperativa <> 1 And vParamAplic.Cooperativa <> 5 Then ' regaixo
            If b Then b = InsertarRecaudacion(cad)
        End If
        
        If b = False Then
            ProcesarFichero = False
            Exit Function
        End If
        
        Line Input #NF, cad
    Wend
    Close #NF
    
    If cad <> "" Then
        '[Monica]09/01/2013: Nueva cooperativa Ribarroja
        If vParamAplic.Cooperativa = 1 Then
            b = InsertarLineaAlz(cad)
        ElseIf vParamAplic.Cooperativa = 5 Then
            b = InsertarLineaRib(cad)
        Else
            b = InsertarLinea(cad)
        End If
'--monica: insertamos recaudacion leyendo de fichero al final del proceso y solo para Alzira
'        If b Then b = InsertarRecaudacion(cad)

'++monica: en regaixo hemos de insertar en srecau para la contabilizacion de turno lo habiamos quitado
'          para alzira
        '[Monica]09/01/2013: nueva cooperativa de Ribarroja
        If vParamAplic.Cooperativa <> 1 And vParamAplic.Cooperativa <> 5 Then ' regaixo
            If b Then b = InsertarRecaudacion(cad)
        End If

        If b = False Then
            ProcesarFichero = False
            Exit Function
        End If
    End If
    
    '++monica: insertamos recaudacion solo para alzira
    '[Monica]09/01/2013: nueva cooperativa de Ribarroja
    If vParamAplic.Cooperativa = 1 Then
        NomFic = LCase(nomFich)
        If b Then b = InsertarRecaudacionAlz(Replace(NomFic, "ventas", "totales"))
    End If
    
    '++monica: insertamos en sturno tanto para alzira como para regaixo
    If vParamAplic.Cooperativa <> 5 Then
        NomFic = LCase(nomFich)
        If b Then b = InsertarLineaTurnoNew(Replace(NomFic, "ventas", "totaliza"))
    End If
    
    ProcesarFichero = b
    
    Pb1.visible = False
    lblProgres(0).Caption = ""
    lblProgres(1).Caption = ""

End Function
                
Private Function ProcesarFichero2(nomFich As String) As Boolean
Dim NF As Long
Dim cad As String
Dim i As Integer
Dim longitud As Long
Dim RS As ADODB.Recordset
Dim Rs1 As ADODB.Recordset
Dim Numreg As Long
Dim Sql As String
Dim SQL1 As String
Dim total As Long
Dim v_cant As Currency
Dim v_impo As Currency
Dim v_prec As Currency
Dim b As Boolean

    On Error GoTo eProcesarFichero2
    
    ProcesarFichero2 = False
    
    NF = FreeFile
    Open nomFich For Input As #NF ' & "\BV" & Format(CDate(txtcodigo(0).Text), "ddmmyy") & "." & Format(txtcodigo(1).Text, "000") For Input As #NF
    
    Line Input #NF, cad
    i = 0
    
    lblProgres(0).Caption = "Insertando en Tabla temporal: " & nomFich
    longitud = FileLen(nomFich)
    
    Pb1.visible = True
    Me.Pb1.Max = longitud
    Me.Refresh
    Me.Pb1.Value = 0
    ' PROCESO DEL FICHERO VENTAS.TXT

    b = True

    While Not EOF(NF) And b
        i = i + 1
        
        Me.Pb1.Value = Me.Pb1.Value + Len(cad)
        lblProgres(1).Caption = "Linea " & i
        Me.Refresh
        '[Monica]09/01/2013: nueva cooperativa de Ribarroja
        If vParamAplic.Cooperativa = 1 Then
            b = ComprobarRegistroAlz(cad)
        ElseIf vParamAplic.Cooperativa = 5 Then
            b = ComprobarRegistroRib(cad)
        Else
            b = ComprobarRegistro(cad)
        End If
        
        Line Input #NF, cad
    Wend
    Close #NF
    
    If cad <> "" Then
        i = i + 1
        
        Me.Pb1.Value = Me.Pb1.Value + Len(cad)
        lblProgres(1).Caption = "Linea " & i
        Me.Refresh
        '[Monica]09/01/2013: nueva cooperativa de Ribarroja
        If vParamAplic.Cooperativa = 1 Then
            b = ComprobarRegistroAlz(cad)
        ElseIf vParamAplic.Cooperativa = 5 Then
            b = ComprobarRegistroRib(cad)
        Else
            b = ComprobarRegistro(cad)
        End If
    End If
    
    Pb1.visible = False
    lblProgres(0).Caption = ""
    lblProgres(1).Caption = ""

    ProcesarFichero2 = b
    Exit Function

eProcesarFichero2:
    ProcesarFichero2 = False
End Function
                
Private Function InsertarCabecera(cad As String) As Boolean
Dim numfactu As String
Dim TipDocu As String
Dim FechaCa As String
Dim Turno As String
Dim Hora As String
Dim forpa As String
Dim Tarje As String
Dim Tarje1 As String
Dim Matric As String
Dim NomCli As String
Dim NifCli As String
Dim Ticket As String
Dim CtaConta As String ' cuenta contable de clientes contado
Dim codsoc As String
Dim Sql As String

    On Error GoTo eInsertarCabecera

    InsertarCabecera = False

    numfactu = 0
    TipDocu = Mid(cad, 10, 1)
    FechaCa = Mid(cad, 11, 2) & Mid(cad, 13, 2) & "20" & Mid(cad, 15, 2)
    Turno = Mid(cad, 17, 1)
    Hora = Mid(cad, 18, 2) & ":" & Mid(cad, 21, 2) & ":00"
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
                
                Sql = "INSERT INTO ssocio (codsocio, codcoope, nomsocio, domsocio, codposta, pobsocio, " & _
                      "prosocio, nifsocio, telsocio, faxsocio, movsocio, maisocio, wwwsocio, fechaalt, " & _
                      "fechabaj, codtarif, codbanco, codsucur, digcontr, cuentaba, impfactu, dtolitro, " & _
                      "codforpa, tipsocio, bonifbas, bonifesp, codsitua, codmacta, obssocio) VALUES (" & _
                      DBSet(Tarje, "N") & ",0," & DBSet(NomCli, "T") & ",'DESCONOCIDA','46','VALENCIA', " & _
                      "'VALENCIA'," & DBSet(NifCli, "T") & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & _
                      DBSet(txtCodigo(0).Text, "F") & "," & _
                      ValorNulo & ",0," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & ",0,0," & _
                      "0,0,0,0,0," & ValorNulo & "," & ValorNulo & ")"
                      
                Conn.Execute Sql
                      
                Sql = "INSERT INTO starje (codsocio, numlinea, numtarje, nomtarje, codbanco, codsucur, " & _
                      "digcontr, cuentaba, tiptarje) VALUES (" & DBSet(Tarje, "N") & ",1," & DBSet(NomCli, "T") & "," & ValorNulo & "," & ValorNulo & "," & _
                      ValorNulo & "," & ValorNulo & ",0)"
                
                Conn.Execute Sql
            End If
    End Select
   

    'MIRAMOS SI EXISTE LA TARJETA
    codsoc = ""
    codsoc = DevuelveDesdeBD("codsocio", "starje", "numtarje", Tarje, "T")
    If Tarje = "       " Then Tarje = "0000000"
    If codsoc = "" Then
    
        Sql = "insert into tmpinformes (codusu, importe1, fecha1, campo1, campo2, importe2, nombre1) values (" & _
              vSesion.Codigo & "," & DBSet(Ticket, "N") & ",'" & Mid(FechaCa, 5, 4) & Mid(FechaCa, 3, 2) & Mid(FechaCa, 1, 2) & "'," & DBSet(Format(Hora, "hh"), "N") & _
              "," & DBSet(Format(Hora, "mm"), "N") & "," & DBSet(Tarje, "N") & ",'Nro. Tarjeta no existe') "
              
        Conn.Execute Sql
        
        
    Else
        Sql = "update scaalb set codsocio = " & DBSet(codsoc, "N") & ", numtarje = " & DBSet(Tarje, "N") & ", numalbar = " & _
               DBSet(Ticket, "T") & ", horalbar = " & DBSet(txtCodigo(0).Text & " " & Hora, "FH") & ", matricul = " & DBSet(Matric, "T") & _
               ", codforpa = " & DBSet(forpa, "N") & ", numfactu = " & DBSet(numfactu, "N") & _
               " where fecalbar = " & DBSet(txtCodigo(0).Text, "F") & " and codturno = " & DBSet(txtCodigo(1).Text, "N") & _
               " and numalbar = " & DBSet(vContad, "T")
               
        Conn.Execute Sql
    End If
    
    vContad = vContad + 1

    InsertarCabecera = True
    
eInsertarCabecera:
    If Err.Number <> 0 Then
        MsgBox "Error en Insertar Cabecera " & Err.Description, vbExclamation
    End If

End Function
            
Private Function ComprobarRegistro(cad As String) As Boolean
Dim Sql As String

Dim Base As String
Dim NombreBase As String
Dim Turno As String
Dim NumAlbaran As String
Dim NumFactura As String
Dim IdVendedor As String
Dim NombreVendedor As String
Dim fechahora As String
Dim CodigoCliente As String
Dim NombreCliente As String
Dim Matricula As String
Dim CodigoProducto As String
Dim Surtidor As String
Dim Manguera As String
Dim PrecioLitro As String
Dim PrecioSinDto As String
Dim cantidad As String
Dim Importe As String
Dim IdTipoPago As String
Dim DescrTipoPago As String
Dim CodigoTipoPago As String
Dim NifCliente As String
Dim IdProducto As String
Dim Tarjeta As String
Dim Tarje As String


Dim c_Cantidad As Currency
Dim c_Importe As Currency
Dim c_Precio As Currency

Dim Fecha As String
Dim Hora As String

Dim Mens As String
Dim Kilometros As String


Dim codsoc As String

    On Error GoTo eComprobarRegistro

    ComprobarRegistro = True

    Base = Mid(cad, 1, 10)
    NombreBase = Mid(cad, 11, 50)
    Turno = Mid(cad, 982, 10) 'txtcodigo(1).Text ' el que yo le diga, antes : Mid(cad, 61, 10)
    If CByte(Turno) > 9 Then Turno = "9"
    
    NumAlbaran = Mid(cad, 72, 19)
    NumFactura = Mid(cad, 94, 17) 'antes 91,20
    IdVendedor = Mid(cad, 121, 10)
    NombreVendedor = Mid(cad, 131, 50)
    fechahora = Mid(cad, 181, 14)
    Fecha = Mid(fechahora, 7, 2) & "/" & Mid(fechahora, 5, 2) & "/" & Mid(fechahora, 1, 4)
    Hora = Mid(fechahora, 9, 6)
    CodigoCliente = Mid(cad, 195, 20)
    NombreCliente = Mid(cad, 215, 70)
    Tarjeta = Mid(cad, 290, 20)
    Matricula = Mid(cad, 370, 20)
    IdProducto = Mid(cad, 493, 20)
    Surtidor = Mid(cad, 538, 10)
    Manguera = Mid(cad, 548, 10)
    
    
    '[Monica]24/08/2015: el precio es sin el descuento en la linea 864, antes ponia 568
    PrecioLitro = Mid(cad, 864, 18)
    
    cantidad = Mid(cad, 650, 18)
    Importe = Mid(cad, 668, 18)
    IdTipoPago = Mid(cad, 784, 10)
    DescrTipoPago = Mid(cad, 794, 25)
    CodigoTipoPago = Mid(cad, 1, 10)
    NifCliente = Mid(cad, 834, 9)
    
    '[Monica]24/06/2013: introducimos los kms em el traspaso
    Kilometros = Mid(cad, 415, 18)
    
    
    If Trim(Importe) = "" Then
        Exit Function
    Else
        If CCur(Importe) = 0 Then Exit Function
    End If
    
    c_Cantidad = Round2(CCur(cantidad) / 100, 2)
    c_Importe = Round2(CCur(Importe) / 100, 2)
    c_Precio = Round2(CCur(PrecioLitro) / 100000, 5)
    
    If Trim(NumFactura) <> "" Then
        'SOLAMENTE EN EL CASO DE QUE SEA FACTURA COMPRUEBO QUE EXISTA EL NIF DEL SOCIO
        'Y SI NO EXISTE INTRODUCIRLO EN LA TABLA DE SOCIOS Y TARJETAS
        Tarje = DevuelveDesdeBDNew(cPTours, "ssocio", "codsocio", "nifsocio", NifCliente, "T")
        If Tarje = "" Then
            Tarje = 900000
            Tarje = SugerirCodigoSiguienteStr("ssocio", "codsocio", "codsocio >= 900000 and codsocio <= 999998")
            
'                CtaConta = ""
'                CtaConta = DevuelveDesdeBD("ctaconta", "sparam", "codparam", "01", "N")
            
            Sql = "INSERT INTO ssocio (codsocio, codcoope, nomsocio, domsocio, codposta, pobsocio, " & _
                  "prosocio, nifsocio, telsocio, faxsocio, movsocio, maisocio, wwwsocio, fechaalt, " & _
                  "fechabaj, codtarif, codbanco, codsucur, digcontr, cuentaba, impfactu, dtolitro, " & _
                  "codforpa, tipsocio, bonifbas, bonifesp, codsitua, codmacta, obssocio) VALUES (" & _
                  DBSet(Tarje, "N") & "," & DBSet(vParamAplic.ColecDefecto, "N") & "," & DBSet(NombreCliente, "T") & ",'DESCONOCIDA','46','VALENCIA', " & _
                  "'VALENCIA'," & DBSet(NifCliente, "T") & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & _
                  DBSet(txtCodigo(0).Text, "F") & "," & _
                  ValorNulo & ",0," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & ",0,0," & _
                  "0,0,0,0,0," & DBSet(vParamAplic.CtaContable, "T") & "," & ValorNulo & ")"
                  
            Conn.Execute Sql
                  
            Sql = "INSERT INTO starje (codsocio, numlinea, numtarje, nomtarje, codbanco, codsucur, " & _
                  "digcontr, cuentaba, tiptarje) VALUES (" & DBSet(Tarje, "N") & ",1," & DBSet(Tarje, "N") & "," & DBSet(NombreCliente, "T") & "," & ValorNulo & "," & ValorNulo & "," & _
                  ValorNulo & "," & ValorNulo & ",0)"
            
            Conn.Execute Sql
        End If
    End If
    
    'MIRAMOS SI EXISTE LA TARJETA
    '[Monica]17/06/2013: añadida la condicion de que la tarjeta no venga con asteriscos: instr(1, Tarjeta, "*") = 0
    If Mid(Tarjeta, 1, 4) <> "****" And Trim(Tarjeta) <> "" And InStr(1, Tarjeta, "*") = 0 Then
        '++monica: 15/02/2008 las tarjetas profesionales tienen 16 caracteres solo analizo los 8 últimos
        If Len(Trim(Tarjeta)) = 16 Then
            Tarjeta = Mid(Tarjeta, 9, 16)
        End If
        '++
        codsoc = ""
        codsoc = DevuelveDesdeBD("codsocio", "starje", "numtarje", Tarjeta, "T")
        If codsoc = "" Then
            Mens = "Nro. Tarjeta no existe"
            Sql = "insert into tmpinformes (codusu, importe1, fecha1, campo1, campo2, importe2, nombre2, importe3, importe4, importe5, nombre1) values (" & _
                  vSesion.Codigo & "," & DBSet(NumAlbaran, "T") & "," & DBSet(Fecha, "F") & "," & DBSet(Mid(Hora, 1, 2), "N") & _
                  "," & DBSet(Mid(Hora, 3, 2), "N") & "," & DBSet(CodigoCliente, "N") & "," & DBSet(Tarjeta, "T") & "," & DBSet(c_Cantidad, "N") & "," & DBSet(c_Precio, "N") & "," & DBSet(c_Importe, "N") & "," & DBSet(Mens, "T") & ")"
                  
            Conn.Execute Sql
            
        End If
    End If
    
    'Comprobamos fechas
    If Not EsFechaOK(Fecha) Then
            Mens = "Fecha incorrecta"
            Sql = "insert into tmpinformes (codusu, importe1, fecha1, campo1, campo2, importe2, nombre2, importe3, " & _
                  "importe4, importe5, nombre1) values (" & _
                  vSesion.Codigo & "," & DBSet(NumAlbaran, "T") & "," & DBSet(Fecha, "F") & "," & DBSet(Mid(Hora, 1, 2), "N")
            Sql = Sql & "," & DBSet(Mid(Hora, 3, 2), "N") & "," & DBSet(CodigoCliente, "N") & "," & DBSet(Fecha, "T") & "," & _
                  DBSet(c_Cantidad, "N") & "," & DBSet(c_Precio, "N") & "," & DBSet(c_Importe, "N") & "," & DBSet(Mens, "T") & ")"
            
            Conn.Execute Sql
    Else
        If CDate(Fecha) <> CDate(txtCodigo(0).Text) Or CByte(Turno) <> CByte(txtCodigo(1).Text) Then
            Mens = "Fecha incorrecta" ' o no es del turno"
            Sql = "insert into tmpinformes (codusu, importe1, fecha1, campo1, campo2, importe2, nombre2, importe3, " & _
                  "importe4, importe5, nombre1) values (" & _
                  vSesion.Codigo & "," & DBSet(NumAlbaran, "T") & "," & DBSet(Fecha, "F") & "," & DBSet(Mid(Hora, 1, 2), "N")
            Sql = Sql & "," & DBSet(Mid(Hora, 3, 2), "N") & "," & DBSet(CodigoCliente, "N") & "," & DBSet(Fecha, "T") & "," & _
                  DBSet(c_Cantidad, "N") & "," & DBSet(c_Precio, "N") & "," & DBSet(c_Importe, "N") & "," & DBSet(Mens, "T") & ")"
            
            Conn.Execute Sql
        End If
    End If
    
    
    'Comprobamos que el articulo existe en sartic
    Sql = ""
    Sql = DevuelveDesdeBDNew(cPTours, "sartic", "codartic", "codartic", IdProducto, "N")
    If Sql = "" Then
        Mens = "No existe el artículo"
        Sql = "insert into tmpinformes (codusu, importe1, fecha1, campo1, campo2, importe2, nombre2, " & _
              "importe3, importe4, importe5, nombre1) values (" & _
              vSesion.Codigo & "," & DBSet(NumAlbaran, "T") & "," & DBSet(Fecha, "F") & "," & DBSet(Mid(Hora, 1, 2), "N")
        Sql = Sql & "," & DBSet(Mid(Hora, 3, 2), "N") & "," & DBSet(CodigoCliente, "N") & "," & DBSet(IdProducto, "T") & "," & _
              DBSet(c_Cantidad, "N") & "," & DBSet(c_Precio, "N") & "," & DBSet(c_Importe, "N") & "," & DBSet(Mens, "T") & ")"
              
        Conn.Execute Sql
    End If
    
    
    'Comprobamos que el socio existe
    If CodigoCliente <> "" Then
        Sql = ""
        Sql = DevuelveDesdeBDNew(cPTours, "ssocio", "codsocio", "codsocio", CodigoCliente, "N")
        If Sql = "" Then
            Mens = "No existe el cliente"
            Sql = "insert into tmpinformes (codusu, importe1, fecha1, campo1, campo2, importe2, nombre2, importe3, " & _
                  "importe4, importe5, nombre1) values (" & _
                  vSesion.Codigo & "," & DBSet(NumAlbaran, "T") & "," & DBSet(Fecha, "F") & "," & DBSet(Mid(Hora, 1, 2), "N")
            Sql = Sql & "," & DBSet(Mid(Hora, 3, 2), "N") & "," & DBSet(CodigoCliente, "N") & "," & DBSet(CodigoCliente, "T") & "," & _
                    DBSet(c_Cantidad, "N") & "," & DBSet(c_Precio, "N") & "," & DBSet(c_Importe, "N") & "," & DBSet(Mens, "T") & ")"
            
            Conn.Execute Sql
        End If
    End If
    
    'Comprobamos que la forma de pago existe
    If IdTipoPago <> "" Then
        Sql = ""
        Sql = DevuelveDesdeBDNew(cPTours, "sforpa", "codforpa", "forpaalvic", IdTipoPago, "N")
        If Sql = "" Then
            Mens = "No existe la forma de pago Alvic"
            Sql = "insert into tmpinformes (codusu, importe1, fecha1, campo1, campo2, importe2, nombre2, " & _
                  "importe3, importe4, importe5, nombre1) values (" & _
                  vSesion.Codigo & "," & DBSet(NumAlbaran, "T") & "," & DBSet(Fecha, "F") & "," & DBSet(Mid(Hora, 1, 2), "N")
            Sql = Sql & "," & DBSet(Mid(Hora, 3, 2), "N") & "," & DBSet(CodigoCliente, "N") & "," & DBSet(IdTipoPago, "T") & "," & _
                    DBSet(c_Cantidad, "N") & "," & DBSet(c_Precio, "N") & "," & DBSet(c_Importe, "N") & "," & DBSet(Mens, "T") & ")"
            
            Conn.Execute Sql
        End If
    End If
    
    'Comprobamos que el codigo de trabajador existe
    Sql = ""
    Sql = DevuelveDesdeBDNew(cPTours, "straba", "codtraba", "codtraba", IdVendedor, "N")
    If Sql = "" Then
        Mens = "No existe el trabajador"
        Sql = "insert into tmpinformes (codusu, importe1, fecha1, campo1, campo2, importe2, nombre2, " & _
              "importe3, importe4, importe5, nombre1) values (" & _
              vSesion.Codigo & "," & DBSet(NumAlbaran, "T") & "," & DBSet(Fecha, "F") & "," & DBSet(Mid(Hora, 1, 2), "N")
        Sql = Sql & "," & DBSet(Mid(Hora, 3, 2), "N") & "," & DBSet(CodigoCliente, "N") & "," & DBSet(IdVendedor, "T") & "," & _
              DBSet(c_Cantidad, "N") & "," & DBSet(c_Precio, "N") & "," & DBSet(c_Importe, "N") & "," & DBSet(Mens, "T") & ")"
              
        Conn.Execute Sql
    End If
eComprobarRegistro:
    If Err.Number <> 0 Then
        ComprobarRegistro = False
    End If
End Function

Private Function ComprobarRegistroAlz(cad As String) As Boolean
Dim Sql As String

Dim Base As String
Dim NombreBase As String
Dim Turno As String
Dim NumAlbaran As String
Dim NumFactura As String
Dim IdVendedor As String
Dim NombreVendedor As String
Dim fechahora As String
Dim CodigoCliente As String
Dim NombreCliente As String
Dim Matricula As String
Dim CodigoProducto As String
Dim Surtidor As String
Dim Manguera As String
Dim PrecioLitro As String
Dim cantidad As String
Dim Importe As String
Dim Descuento As String
Dim IdTipoPago As String
Dim DescrTipoPago As String
Dim CodigoTipoPago As String
Dim NifCliente As String
Dim IdProducto As String
Dim Tarjeta As String
Dim Tarje As String


Dim c_Cantidad As Currency
Dim c_Importe As Currency
Dim c_Importe1 As Currency
Dim c_Importe2 As Currency
Dim c_Precio As Currency
Dim c_Precio2 As Currency
Dim c_Descuento As Currency

Dim Fecha As String
Dim Hora As String

Dim Mens As String
Dim Kilometros As String

Dim codsoc As String

Dim IvaArticulo As String
Dim NombreArticulo As String
Dim NomArtic As String
Dim CodIVA As String
Dim PorcIva As Currency

    On Error GoTo eComprobarRegistroAlz

    ComprobarRegistroAlz = True

    Base = Mid(cad, 1, 10)
    NombreBase = Mid(cad, 11, 50)
    Turno = Mid(cad, 982, 10) 'txtcodigo(1).Text ' el que yo le diga, antes : Mid(cad, 61, 10)
    If CByte(Turno) > 9 Then Turno = "9"
    
    NumAlbaran = Mid(cad, 71, 20)
    NumFactura = Mid(cad, 94, 17) 'antes 91,20
    IdVendedor = Mid(cad, 121, 10)
    NombreVendedor = Mid(cad, 131, 50)
    fechahora = Mid(cad, 181, 14)
    Fecha = Mid(fechahora, 7, 2) & "/" & Mid(fechahora, 5, 2) & "/" & Mid(fechahora, 1, 4)
    Hora = Mid(fechahora, 9, 6)
'    CodigoCliente = Mid(cad, 195, 20)
    NombreCliente = Mid(cad, 215, 70)
'    Tarjeta = Mid(cad, 290, 20)
    Tarjeta = Mid(cad, 195, 20)
    Matricula = Mid(cad, 370, 20)
    IdProducto = Mid(cad, 493, 20)
    Surtidor = Mid(cad, 538, 10)
    Manguera = Mid(cad, 548, 10)

    PrecioLitro = Mid(cad, 568, 18)
    
    '[Monica]29/10/2015:faltaban ¿?¿?
    cantidad = Mid(cad, 650, 18)
    Importe = Mid(cad, 668, 18)
    
    
    
    
    Descuento = Mid(cad, 586, 18)
    IdTipoPago = Mid(cad, 784, 10)
    DescrTipoPago = Mid(cad, 794, 25)
    CodigoTipoPago = Mid(cad, 1, 10)
    NifCliente = Mid(cad, 834, 9)
    
    '[Monica]25/06/2013: comprobamos el nombre del articulo y el iva
    '                    en el caso de que el nombre del articulo no coincida mostramos informe pero dejamos continuar
    '                    en el caso de que no coincida el iva mostramos informe y NO dejamos continuar
    IvaArticulo = Mid(cad, 609, 5) ' 5 posiciones 2 decimales implicitos
    NombreArticulo = Mid(cad, 513, 25) ' nombre del articulo para comprobarlo
    
    
    '[Monica]24/06/2013: introducimos los kms em el traspaso
    Kilometros = Mid(cad, 415, 18)
    
    
    If Trim(Importe) = "" Then
        Exit Function
    Else
        If CCur(Importe) = 0 Then Exit Function
    End If
    
    c_Cantidad = Round2(CCur(cantidad) / 100, 2)
    c_Importe = Round2(CCur(Importe) / 100, 2)
    c_Precio = Round2(CCur(PrecioLitro) / 100000, 5)
    
    If Trim(Descuento) <> "" Then
        If CCur(Descuento) <> 0 Then
            c_Descuento = Round2(CCur(Descuento) / 100000, 5)
            c_Importe1 = Round2(c_Cantidad * c_Precio, 2)
            c_Importe2 = c_Importe - c_Importe1
            c_Importe = c_Importe1
            c_Precio2 = Round2(c_Importe2 / c_Cantidad * (-1), 3)
        Else
            c_Descuento = 0
        End If
    End If
    
'[Monica]30/11/2010: añadida la segunda condicion
    If Trim(NumFactura) <> "" And InStr(1, Tarjeta, "Z") <> 0 Then
        'SOLAMENTE EN EL CASO DE QUE SEA FACTURA COMPRUEBO QUE EXISTA EL NIF DEL SOCIO
        'Y SI NO EXISTE INTRODUCIRLO EN LA TABLA DE SOCIOS Y TARJETAS
        Tarje = DevuelveDesdeBDNew(cPTours, "ssocio", "codsocio", "nifsocio", NifCliente, "T")
        If Tarje = "" Then
            Tarje = 900000
            Tarje = SugerirCodigoSiguienteStr("ssocio", "codsocio", "codsocio >= 900000 and codsocio <= 999998")
            
'                CtaConta = ""
'                CtaConta = DevuelveDesdeBD("ctaconta", "sparam", "codparam", "01", "N")
            
            Sql = "INSERT INTO ssocio (codsocio, codcoope, nomsocio, domsocio, codposta, pobsocio, " & _
                  "prosocio, nifsocio, telsocio, faxsocio, movsocio, maisocio, wwwsocio, fechaalt, " & _
                  "fechabaj, codtarif, codbanco, codsucur, digcontr, cuentaba, impfactu, dtolitro, " & _
                  "codforpa, tipsocio, bonifbas, bonifesp, codsitua, codmacta, obssocio) VALUES (" & _
                  DBSet(Tarje, "N") & "," & DBSet(vParamAplic.ColecDefecto, "N") & "," & DBSet(NombreCliente, "T") & ",'DESCONOCIDA','46','VALENCIA', " & _
                  "'VALENCIA'," & DBSet(NifCliente, "T") & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & _
                  DBSet(txtCodigo(0).Text, "F") & "," & _
                  ValorNulo & ",0," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & ",0,0," & _
                  "0,0,0,0,0," & DBSet(vParamAplic.CtaContable, "T") & "," & ValorNulo & ")"
                  
            Conn.Execute Sql
                  
            Sql = "INSERT INTO starje (codsocio, numlinea, numtarje, nomtarje, codbanco, codsucur, " & _
                  "digcontr, cuentaba, tiptarje) VALUES (" & DBSet(Tarje, "N") & ",1," & DBSet(Tarje, "N") & "," & DBSet(NombreCliente, "T") & "," & ValorNulo & "," & ValorNulo & "," & _
                  ValorNulo & "," & ValorNulo & ",0)"
            
            Conn.Execute Sql
            
        Else '[Monica]07/02/2011: caso de que sea un socio que quiere la factura ( me viene en fichero nro de factura y Z )
             ' añadida esta parte del else que no estaba
            If CLng(Tarje) >= 900000 Then
                ' miro si existe tarjeta sino la creo
                Sql = "select count(*) from starje where codsocio= " & DBSet(Tarje, "N")
                If TotalRegistros(Sql) = 0 Then
                    Sql = "INSERT INTO starje (codsocio, numlinea, numtarje, nomtarje, codbanco, codsucur, " & _
                          "digcontr, cuentaba, tiptarje) VALUES (" & DBSet(Tarje, "N") & ",1," & DBSet(Tarje, "N") & "," & DBSet(NombreCliente, "T") & "," & ValorNulo & "," & ValorNulo & "," & _
                          ValorNulo & "," & ValorNulo & ",0)"
                    
                    Conn.Execute Sql
                End If
            Else
                ' el socio es inferior a 900000 miro si hay tarjeta dependiendo del producto
                Dim TipArtic As Integer
                TipArtic = DevuelveValor("select tipogaso from sartic where codartic = " & DBSet(IdProducto, "N"))
                If TipArtic = 3 Then ' si el articulo es gasoleo bonificado
                    Sql = "select count(*) from starje where codsocio= " & DBSet(Tarje, "N") & " and tiptarje = 1"
                    If TotalRegistros(Sql) = 0 Then
                        
                        '[Monica]22/11/2011: si es un cliente de paso y no tiene tarjeta de bonificado le ponemos la que tenga
                        ' he quitado el control de que no existe tarjeta de gasoleo bonificado
'                        Mens = "Tarjeta bonif.no existe"
'                        SQL = "insert into tmpinformes (codusu, importe1, fecha1, campo1, campo2, importe2, nombre2, importe3, importe4, importe5, nombre1) values (" & _
'                              vSesion.Codigo & "," & DBSet(NumAlbaran, "T") & "," & DBSet(Fecha, "F") & "," & DBSet(Mid(Hora, 1, 2), "N") & _
'                              "," & DBSet(Mid(Hora, 3, 2), "N") & "," & DBSet(Tarje, "N") & "," & DBSet(Tarje, "T") & "," & DBSet(c_Cantidad, "N") & "," & DBSet(c_Precio, "N") & "," & DBSet(c_Importe, "N") & "," & DBSet(Mens, "T") & ")"
'
'                        Conn.Execute SQL
                        
                        '[Monica]22/11/2011: si es un cliente de paso y no tiene tarjeta de bonificado le ponemos la que tenga
                        '                    le pongo la primera tarjeta que exista o se la creo si no existe ninguna
                        Sql = "select count(*) from starje where codsocio= " & DBSet(Tarje, "N")
                        If TotalRegistros(Sql) = 0 Then
                            Sql = "INSERT INTO starje (codsocio, numlinea, numtarje, nomtarje, codbanco, codsucur, " & _
                                  "digcontr, cuentaba, tiptarje) VALUES (" & DBSet(Tarje, "N") & ",1," & DBSet(Tarje, "N") & "," & DBSet(NombreCliente, "T") & "," & ValorNulo & "," & ValorNulo & "," & _
                                  ValorNulo & "," & ValorNulo & ",0)"

                            Conn.Execute Sql
                        End If

                    End If
                Else
                    Sql = "select count(*) from starje where codsocio= " & DBSet(Tarje, "N") & " and tiptarje = 0"
                    If TotalRegistros(Sql) = 0 Then
                        Mens = "Nro. Tarjeta no existe"
                        Sql = "insert into tmpinformes (codusu, importe1, fecha1, campo1, campo2, importe2, nombre2, importe3, importe4, importe5, nombre1) values (" & _
                              vSesion.Codigo & "," & DBSet(NumAlbaran, "T") & "," & DBSet(Fecha, "F") & "," & DBSet(Mid(Hora, 1, 2), "N") & _
                              "," & DBSet(Mid(Hora, 3, 2), "N") & "," & DBSet(Tarje, "N") & "," & DBSet(Tarje, "T") & "," & DBSet(c_Cantidad, "N") & "," & DBSet(c_Precio, "N") & "," & DBSet(c_Importe, "N") & "," & DBSet(Mens, "T") & ")"
                        
                        Conn.Execute Sql
                    End If
                End If
            End If '07/02/2011: hasta aqui la parte añadida
        
        End If
    Else
        'MIRAMOS SI EXISTE LA TARJETA
        ' en alzira lo pongo dentro
        codsoc = ""
        
        '[Monica]25/06/2013: el importeb1 ponemos si dejamos o no continuar cuando es 0 no dejamos continuar con 1 sí
        
        '++monica:050508 el numero de tarjeta puede venir a blanco--> dar error
        If Trim(Tarjeta) <> "" Then codsoc = DevuelveDesdeBD("codsocio", "starje", "numtarje", Tarjeta, "N")
        If codsoc = "" Then
            Mens = "Nro. Tarjeta no existe"
            Sql = "insert into tmpinformes (codusu, importe1, fecha1, campo1, campo2, importe2, nombre2, importe3, importe4, importe5, nombre1) values (" & _
                  vSesion.Codigo & "," & DBSet(NumAlbaran, "T") & "," & DBSet(Fecha, "F") & "," & DBSet(Mid(Hora, 1, 2), "N") & _
                  "," & DBSet(Mid(Hora, 3, 2), "N") & "," & DBSet(CodigoCliente, "N") & "," & DBSet(Tarjeta, "T") & "," & DBSet(c_Cantidad, "N") & "," & DBSet(c_Precio, "N") & "," & DBSet(c_Importe, "N") & "," & DBSet(Mens, "T") & ")"
                  
            Conn.Execute Sql
        Else
            ' comprobamos que el socio existe
            ' no haria falta pq hay clave referencial a ssocio
            Sql = ""
            Sql = DevuelveDesdeBDNew(cPTours, "ssocio", "codsocio", "codsocio", codsoc, "N")
            If Sql = "" Then
                Mens = "No existe el cliente"
                Sql = "insert into tmpinformes (codusu, importe1, fecha1, campo1, campo2, importe2, nombre2, importe3, " & _
                      "importe4, importe5, nombre1) values (" & _
                      vSesion.Codigo & "," & DBSet(NumAlbaran, "T") & "," & DBSet(Fecha, "F") & "," & DBSet(Mid(Hora, 1, 2), "N")
                Sql = Sql & "," & DBSet(Mid(Hora, 3, 2), "N") & "," & DBSet(codsoc, "N") & "," & DBSet(codsoc, "T") & "," & _
                        DBSet(c_Cantidad, "N") & "," & DBSet(c_Precio, "N") & "," & DBSet(c_Importe, "N") & "," & DBSet(Mens, "T") & ")"
                
                Conn.Execute Sql
            End If
        End If
    End If
    
'    'MIRAMOS SI EXISTE LA TARJETA
'    If Mid(Tarjeta, 1, 4) <> "****" And Trim(Tarjeta) <> "" Then
''        '++monica: 15/02/2008 las tarjetas profesionales tienen 16 caracteres solo analizo los 8 últimos
''        If Len(Trim(Tarjeta)) = 16 Then
''            Tarjeta = Mid(Tarjeta, 9, 16)
''        End If
''        '++
'        codsoc = ""
'        codsoc = DevuelveDesdeBD("codsocio", "starje", "numtarje", Tarjeta, "T")
'        If codsoc = "" Then
'            Mens = "Nro. Tarjeta no existe"
'            Sql = "insert into tmpinformes (codusu, importe1, fecha1, campo1, campo2, importe2, nombre2, importe3, importe4, importe5, nombre1) values (" & _
'                  vSesion.Codigo & "," & DBSet(NumAlbaran, "T") & "," & DBSet(Fecha, "F") & "," & DBSet(Mid(Hora, 1, 2), "N") & _
'                  "," & DBSet(Mid(Hora, 3, 2), "N") & "," & DBSet(CodigoCliente, "N") & "," & DBSet(Tarjeta, "T") & "," & DBSet(c_Cantidad, "N") & "," & DBSet(c_Precio, "N") & "," & DBSet(c_Importe, "N") & "," & DBSet(Mens, "T") & ")"
'
'            Conn.Execute Sql
'        Else
'            ' comprobamos que el socio existe
'            ' no haria falta pq hay clave referencial a ssocio
'            Sql = ""
'            Sql = DevuelveDesdeBDNew(cPTours, "ssocio", "codsocio", "codsocio", codsoc, "N")
'            If Sql = "" Then
'                Mens = "No existe el cliente"
'                Sql = "insert into tmpinformes (codusu, importe1, fecha1, campo1, campo2, importe2, nombre2, importe3, " & _
'                      "importe4, importe5, nombre1) values (" & _
'                      vSesion.Codigo & "," & DBSet(NumAlbaran, "T") & "," & DBSet(Fecha, "F") & "," & DBSet(Mid(Hora, 1, 2), "N")
'                Sql = Sql & "," & DBSet(Mid(Hora, 3, 2), "N") & "," & DBSet(codsoc, "N") & "," & DBSet(codsoc, "T") & "," & _
'                        DBSet(c_Cantidad, "N") & "," & DBSet(c_Precio, "N") & "," & DBSet(c_Importe, "N") & "," & DBSet(Mens, "T") & ")"
'
'                Conn.Execute Sql
'            End If
'
'        End If
'    End If
    
    'Comprobamos fechas
    If Not EsFechaOK(Fecha) Then
            Mens = "Fecha incorrecta"
            Sql = "insert into tmpinformes (codusu, importe1, fecha1, campo1, campo2, importe2, nombre2, importe3, " & _
                  "importe4, importe5, nombre1) values (" & _
                  vSesion.Codigo & "," & DBSet(NumAlbaran, "T") & "," & DBSet(Fecha, "F") & "," & DBSet(Mid(Hora, 1, 2), "N")
            Sql = Sql & "," & DBSet(Mid(Hora, 3, 2), "N") & "," & DBSet(CodigoCliente, "N") & "," & DBSet(Fecha, "T") & "," & _
                  DBSet(c_Cantidad, "N") & "," & DBSet(c_Precio, "N") & "," & DBSet(c_Importe, "N") & "," & DBSet(Mens, "T") & ")"
            
            Conn.Execute Sql
    Else
        If CDate(Fecha) <> CDate(txtCodigo(0).Text) Or CByte(Turno) <> CByte(txtCodigo(1).Text) Then
            Mens = "Fecha incorrecta" ' o no es del turno"
            Sql = "insert into tmpinformes (codusu, importe1, fecha1, campo1, campo2, importe2, nombre2, importe3, " & _
                  "importe4, importe5, nombre1) values (" & _
                  vSesion.Codigo & "," & DBSet(NumAlbaran, "T") & "," & DBSet(Fecha, "F") & "," & DBSet(Mid(Hora, 1, 2), "N")
            Sql = Sql & "," & DBSet(Mid(Hora, 3, 2), "N") & "," & DBSet(CodigoCliente, "N") & "," & DBSet(Fecha, "T") & "," & _
                  DBSet(c_Cantidad, "N") & "," & DBSet(c_Precio, "N") & "," & DBSet(c_Importe, "N") & "," & DBSet(Mens, "T") & ")"
            
            Conn.Execute Sql
        End If
    End If
    
    
    'Comprobamos que el articulo existe en sartic
    NomArtic = "nomartic"
    
    Sql = ""
    Sql = DevuelveDesdeBDNew(cPTours, "sartic", "codartic", "codartic", IdProducto, "N", NomArtic)
    If Sql = "" Then
        Mens = "No existe el artículo"
        Dim IdProducto1 As Currency
        IdProducto1 = CCur(IdProducto)
        Sql = "insert into tmpinformes (codusu, importe1, fecha1, campo1, campo2, importe2, nombre2, " & _
              "importe3, importe4, importe5, nombre1) values (" & _
              vSesion.Codigo & "," & DBSet(NumAlbaran, "T") & "," & DBSet(Fecha, "F") & "," & DBSet(Mid(Hora, 1, 2), "N")
        Sql = Sql & "," & DBSet(Mid(Hora, 3, 2), "N") & "," & DBSet(CodigoCliente, "N") & "," & DBSet(IdProducto1, "T") & "," & _
              DBSet(c_Cantidad, "N") & "," & DBSet(c_Precio, "N") & "," & DBSet(c_Importe, "N") & "," & DBSet(Mens, "T") & ")"
              
        Conn.Execute Sql
    Else
        '[Monica]25/06/2013: añadimos el else de si el nombre es distinto y es distinto iva
        If Trim(NomArtic) <> Trim(NombreArticulo) And Not EsArticuloCombustible(IdProducto) Then
            Mens = "Nombre art." & Format(IdProducto, "000000") & " no coincide"
            Sql = "insert into tmpinformes (codusu, importe1, fecha1, campo1, campo2, importe2, nombre2, " & _
                  "importe3, importe4, importe5, nombre1, importeb1) values (" & _
                  vSesion.Codigo & "," & DBSet(NumAlbaran, "T") & "," & DBSet(Fecha, "F") & "," & DBSet(Mid(Hora, 1, 2), "N")
            Sql = Sql & "," & DBSet(Mid(Hora, 3, 2), "N") & "," & DBSet(CodigoCliente, "N") & "," & DBSet(NomArtic, "T") & "," & _
                  DBSet(c_Cantidad, "N") & "," & DBSet(c_Precio, "N") & "," & DBSet(c_Importe, "N") & "," & DBSet(Mens, "T") & ",0)"
                  
            Conn.Execute Sql
        End If
        
        CodIVA = ""
        CodIVA = DevuelveDesdeBDNew(cPTours, "sartic", "codigiva", "codartic", IdProducto, "N")
        PorcIva = 0
        If CodIVA <> "" Then
            PorcIva = DevuelveDesdeBDNew(cConta, "tiposiva", "porceiva", "codigiva", CodIVA, "N")
        End If
            
        ' aquí no dejamos continuar
        If PorcIva <> Round2(CInt(ComprobarCero(IvaArticulo)) / 100, 0) Then
            Mens = "Porcentaje de iva distinto"
            Sql = "insert into tmpinformes (codusu, importe1, fecha1, campo1, campo2, importe2, nombre2, " & _
                  "importe3, importe4, importe5, nombre1) values (" & _
                  vSesion.Codigo & "," & DBSet(NumAlbaran, "T") & "," & DBSet(Fecha, "F") & "," & DBSet(Mid(Hora, 1, 2), "N")
            Sql = Sql & "," & DBSet(Mid(Hora, 3, 2), "N") & "," & DBSet(CodigoCliente, "N") & "," & DBSet(IdProducto, "T") & "," & _
                  DBSet(c_Cantidad, "N") & "," & DBSet(c_Precio, "N") & "," & DBSet(c_Importe, "N") & "," & DBSet(Mens, "T") & ")"
                  
            Conn.Execute Sql
        End If
    End If
    
    
'    'Comprobamos que el socio existe
'    If CodigoCliente <> "" Then
'        Sql = ""
'        Sql = DevuelveDesdeBDNew(cPTours, "ssocio", "codsocio", "codsocio", CodigoCliente, "N")
'        If Sql = "" Then
'            Mens = "No existe el cliente"
'            Sql = "insert into tmpinformes (codusu, importe1, fecha1, campo1, campo2, importe2, nombre2, importe3, " & _
'                  "importe4, importe5, nombre1) values (" & _
'                  vSesion.Codigo & "," & DBSet(NumAlbaran, "T") & "," & DBSet(Fecha, "F") & "," & DBSet(Mid(Hora, 1, 2), "N")
'            Sql = Sql & "," & DBSet(Mid(Hora, 3, 2), "N") & "," & DBSet(CodigoCliente, "N") & "," & DBSet(CodigoCliente, "T") & "," & _
'                    DBSet(c_Cantidad, "N") & "," & DBSet(c_Precio, "N") & "," & DBSet(c_Importe, "N") & "," & DBSet(Mens, "T") & ")"
'
'            Conn.Execute Sql
'        End If
'    End If
    
    'Comprobamos que la forma de pago existe
    If IdTipoPago <> "" Then
        Sql = ""
        Sql = DevuelveDesdeBDNew(cPTours, "sforpa", "codforpa", "forpaalvic", IdTipoPago, "N")
        If Sql = "" Then
            Mens = "No existe la forma de pago Alvic"
            Sql = "insert into tmpinformes (codusu, importe1, fecha1, campo1, campo2, importe2, nombre2, " & _
                  "importe3, importe4, importe5, nombre1) values (" & _
                  vSesion.Codigo & "," & DBSet(NumAlbaran, "T") & "," & DBSet(Fecha, "F") & "," & DBSet(Mid(Hora, 1, 2), "N")
            Sql = Sql & "," & DBSet(Mid(Hora, 3, 2), "N") & "," & DBSet(CodigoCliente, "N") & "," & DBSet(IdTipoPago, "T") & "," & _
                    DBSet(c_Cantidad, "N") & "," & DBSet(c_Precio, "N") & "," & DBSet(c_Importe, "N") & "," & DBSet(Mens, "T") & ")"
            
            Conn.Execute Sql
        End If
    End If
    
    'Comprobamos que el codigo de trabajador existe
    Sql = ""
    Sql = DevuelveDesdeBDNew(cPTours, "straba", "codtraba", "codtraba", IdVendedor, "N")
    If Sql = "" Then
        Mens = "No existe el trabajador"
        Sql = "insert into tmpinformes (codusu, importe1, fecha1, campo1, campo2, importe2, nombre2, " & _
              "importe3, importe4, importe5, nombre1) values (" & _
              vSesion.Codigo & "," & DBSet(NumAlbaran, "T") & "," & DBSet(Fecha, "F") & "," & DBSet(Mid(Hora, 1, 2), "N")
        Sql = Sql & "," & DBSet(Mid(Hora, 3, 2), "N") & "," & DBSet(CodigoCliente, "N") & "," & DBSet(IdVendedor, "T") & "," & _
              DBSet(c_Cantidad, "N") & "," & DBSet(c_Precio, "N") & "," & DBSet(c_Importe, "N") & "," & DBSet(Mens, "T") & ")"
              
        Conn.Execute Sql
    End If
    
    'Comprobamos si hay descuento que el codigo de articulo de dto existe
    If c_Descuento <> 0 Then
        Sql = ""
        Sql = DevuelveDesdeBDNew(cPTours, "sartic", "artdto", "codartic", IdProducto, "N")
        If Sql = "" Then
            Mens = "No tiene artículo de descuento"
            Sql = "insert into tmpinformes (codusu, importe1, fecha1, campo1, campo2, importe2, nombre2, " & _
                  "importe3, importe4, importe5, nombre1) values (" & _
                  vSesion.Codigo & "," & DBSet(NumAlbaran, "T") & "," & DBSet(Fecha, "F") & "," & DBSet(Mid(Hora, 1, 2), "N")
            Sql = Sql & "," & DBSet(Mid(Hora, 3, 2), "N") & "," & DBSet(CodigoCliente, "N") & "," & DBSet(IdProducto, "T") & "," & _
                  DBSet(c_Cantidad, "N") & "," & DBSet(c_Precio, "N") & "," & DBSet(c_Importe, "N") & "," & DBSet(Mens, "T") & ")"
                  
            Conn.Execute Sql
        End If
    End If
    
eComprobarRegistroAlz:
    If Err.Number <> 0 Then
        ComprobarRegistroAlz = False
    End If
End Function
            
            
Private Function ComprobarRegistroRib(cad As String) As Boolean
Dim Sql As String

Dim Base As String
Dim NombreBase As String
Dim Turno As String
Dim NumAlbaran As String
Dim NumFactura As String
Dim IdVendedor As String
Dim NombreVendedor As String
Dim fechahora As String
Dim CodigoCliente As String
Dim NombreCliente As String
Dim Matricula As String
Dim CodigoProducto As String
Dim Surtidor As String
Dim Manguera As String
Dim PrecioLitro As String
Dim cantidad As String
Dim Importe As String
Dim Descuento As String
Dim IdTipoPago As String
Dim DescrTipoPago As String
Dim CodigoTipoPago As String
Dim NifCliente As String
Dim IdProducto As String
Dim Tarjeta As String
Dim Tarje As String


Dim c_Cantidad As Currency
Dim c_Importe As Currency
Dim c_Importe1 As Currency
Dim c_Importe2 As Currency
Dim c_Precio As Currency
Dim c_Precio2 As Currency
Dim c_Descuento As Currency

Dim Fecha As String
Dim Hora As String

Dim Mens As String
Dim Kilometros As String

Dim codsoc As String

    On Error GoTo eComprobarRegistroRib

    ComprobarRegistroRib = True

    Base = Mid(cad, 1, 10)
    NombreBase = Mid(cad, 11, 50)
    Turno = Mid(cad, 982, 10) 'txtcodigo(1).Text ' el que yo le diga, antes : Mid(cad, 61, 10)
    If CByte(Turno) > 9 Then Turno = "9"
    
    NumAlbaran = Mid(cad, 71, 20)
    NumFactura = Mid(cad, 92, 7) 'antes 91,20
    IdVendedor = Mid(cad, 121, 10)
    NombreVendedor = Mid(cad, 131, 50)
    fechahora = Mid(cad, 181, 14)
    Fecha = Mid(fechahora, 7, 2) & "/" & Mid(fechahora, 5, 2) & "/" & Mid(fechahora, 1, 4)
    Hora = Mid(fechahora, 9, 6)
'    CodigoCliente = Mid(cad, 195, 20)
    NombreCliente = Mid(cad, 215, 70)
    Tarjeta = Mid(cad, 195, 20)
    Matricula = Mid(cad, 370, 20)
    IdProducto = Mid(cad, 493, 20)
    Surtidor = Mid(cad, 538, 10)
    Manguera = Mid(cad, 548, 10)
    PrecioLitro = Mid(cad, 568, 18)
    cantidad = Mid(cad, 650, 18)
    Importe = Mid(cad, 668, 18)
    Descuento = Mid(cad, 586, 18)
    IdTipoPago = Mid(cad, 784, 10)
    DescrTipoPago = Mid(cad, 794, 25)
    CodigoTipoPago = Mid(cad, 1, 10)
    NifCliente = Mid(cad, 834, 9)
    
    '[Monica]24/06/2013: introducimos los kms em el traspaso
    Kilometros = Mid(cad, 415, 18)
    
    
    
    If Trim(Importe) = "" Then
        Exit Function
    Else
        If CCur(Importe) = 0 Then Exit Function
    End If
    
    c_Cantidad = Round2(CCur(cantidad) / 100, 2)
    c_Importe = Round2(CCur(Importe) / 100, 2)
    c_Precio = Round2(CCur(PrecioLitro) / 100000, 5)
    
    If Trim(Descuento) <> "" Then
        If CCur(Descuento) <> 0 Then
            c_Descuento = Round2(CCur(Descuento) / 100000, 5)
            c_Importe1 = Round2(c_Cantidad * c_Precio, 2)
            c_Importe2 = c_Importe - c_Importe1
            c_Importe = c_Importe1
            c_Precio2 = Round2(c_Importe2 / c_Cantidad * (-1), 3)
        Else
            c_Descuento = 0
        End If
    End If
    
'[Monica]30/11/2010: añadida la segunda condicion
    If Trim(NumFactura) <> "" And InStr(1, Tarjeta, "Z") <> 0 Then
        'SOLAMENTE EN EL CASO DE QUE SEA FACTURA COMPRUEBO QUE EXISTA EL NIF DEL SOCIO
        'Y SI NO EXISTE INTRODUCIRLO EN LA TABLA DE SOCIOS Y TARJETAS
        Tarje = DevuelveDesdeBDNew(cPTours, "ssocio", "codsocio", "nifsocio", NifCliente, "T")
        If Tarje = "" Then
            Tarje = 900000
            Tarje = SugerirCodigoSiguienteStr("ssocio", "codsocio", "codsocio >= 900000 and codsocio <= 999998")
            
'                CtaConta = ""
'                CtaConta = DevuelveDesdeBD("ctaconta", "sparam", "codparam", "01", "N")
            
            Sql = "INSERT INTO ssocio (codsocio, codcoope, nomsocio, domsocio, codposta, pobsocio, " & _
                  "prosocio, nifsocio, telsocio, faxsocio, movsocio, maisocio, wwwsocio, fechaalt, " & _
                  "fechabaj, codtarif, codbanco, codsucur, digcontr, cuentaba, impfactu, dtolitro, " & _
                  "codforpa, tipsocio, bonifbas, bonifesp, codsitua, codmacta, obssocio) VALUES (" & _
                  DBSet(Tarje, "N") & "," & DBSet(vParamAplic.ColecDefecto, "N") & "," & DBSet(NombreCliente, "T") & ",'DESCONOCIDA','46','VALENCIA', " & _
                  "'VALENCIA'," & DBSet(NifCliente, "T") & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & _
                  DBSet(txtCodigo(0).Text, "F") & "," & _
                  ValorNulo & ",0," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & ",0,0," & _
                  "1,0,0,0,0," & DBSet(vParamAplic.CtaContable, "T") & "," & ValorNulo & ")"
                  
            Conn.Execute Sql
                  
            Sql = "INSERT INTO starje (codsocio, numlinea, numtarje, nomtarje, codbanco, codsucur, " & _
                  "digcontr, cuentaba, tiptarje) VALUES (" & DBSet(Tarje, "N") & ",1," & DBSet(Tarje, "N") & "," & DBSet(NombreCliente, "T") & "," & ValorNulo & "," & ValorNulo & "," & _
                  ValorNulo & "," & ValorNulo & ",0)"
            
            Conn.Execute Sql
            
        Else '[Monica]07/02/2011: caso de que sea un socio que quiere la factura ( me viene en fichero nro de factura y Z )
             ' añadida esta parte del else que no estaba
            If CLng(Tarje) >= 900000 Then
                ' miro si existe tarjeta sino la creo
                Sql = "select count(*) from starje where codsocio= " & DBSet(Tarje, "N")
                If TotalRegistros(Sql) = 0 Then
                    Sql = "INSERT INTO starje (codsocio, numlinea, numtarje, nomtarje, codbanco, codsucur, " & _
                          "digcontr, cuentaba, tiptarje) VALUES (" & DBSet(Tarje, "N") & ",1," & DBSet(Tarje, "N") & "," & DBSet(NombreCliente, "T") & "," & ValorNulo & "," & ValorNulo & "," & _
                          ValorNulo & "," & ValorNulo & ",0)"
                    
                    Conn.Execute Sql
                End If
            Else
                ' el socio es inferior a 900000 miro si hay tarjeta dependiendo del producto
                Dim TipArtic As Integer
                TipArtic = DevuelveValor("select tipogaso from sartic where codartic = " & DBSet(IdProducto, "N"))
                If TipArtic = 3 Then ' si el articulo es gasoleo bonificado
                    Sql = "select count(*) from starje where codsocio= " & DBSet(Tarje, "N") & " and tiptarje = 1"
                    If TotalRegistros(Sql) = 0 Then
                        
                        '[Monica]22/11/2011: si es un cliente de paso y no tiene tarjeta de bonificado le ponemos la que tenga
                        ' he quitado el control de que no existe tarjeta de gasoleo bonificado
'                        Mens = "Tarjeta bonif.no existe"
'                        SQL = "insert into tmpinformes (codusu, importe1, fecha1, campo1, campo2, importe2, nombre2, importe3, importe4, importe5, nombre1) values (" & _
'                              vSesion.Codigo & "," & DBSet(NumAlbaran, "T") & "," & DBSet(Fecha, "F") & "," & DBSet(Mid(Hora, 1, 2), "N") & _
'                              "," & DBSet(Mid(Hora, 3, 2), "N") & "," & DBSet(Tarje, "N") & "," & DBSet(Tarje, "T") & "," & DBSet(c_Cantidad, "N") & "," & DBSet(c_Precio, "N") & "," & DBSet(c_Importe, "N") & "," & DBSet(Mens, "T") & ")"
'
'                        Conn.Execute SQL
                        
                        '[Monica]22/11/2011: si es un cliente de paso y no tiene tarjeta de bonificado le ponemos la que tenga
                        '                    le pongo la primera tarjeta que exista o se la creo si no existe ninguna
                        Sql = "select count(*) from starje where codsocio= " & DBSet(Tarje, "N")
                        If TotalRegistros(Sql) = 0 Then
                            Sql = "INSERT INTO starje (codsocio, numlinea, numtarje, nomtarje, codbanco, codsucur, " & _
                                  "digcontr, cuentaba, tiptarje) VALUES (" & DBSet(Tarje, "N") & ",1," & DBSet(Tarje, "N") & "," & DBSet(NombreCliente, "T") & "," & ValorNulo & "," & ValorNulo & "," & _
                                  ValorNulo & "," & ValorNulo & ",0)"

                            Conn.Execute Sql
                        End If

                    End If
                Else
                    Sql = "select count(*) from starje where codsocio= " & DBSet(Tarje, "N") & " and tiptarje = 0"
                    If TotalRegistros(Sql) = 0 Then
                        Mens = "Nro. Tarjeta no existe"
                        Sql = "insert into tmpinformes (codusu, importe1, fecha1, campo1, campo2, importe2, nombre2, importe3, importe4, importe5, nombre1) values (" & _
                              vSesion.Codigo & "," & DBSet(NumAlbaran, "T") & "," & DBSet(Fecha, "F") & "," & DBSet(Mid(Hora, 1, 2), "N") & _
                              "," & DBSet(Mid(Hora, 3, 2), "N") & "," & DBSet(Tarje, "N") & "," & DBSet(Tarje, "T") & "," & DBSet(c_Cantidad, "N") & "," & DBSet(c_Precio, "N") & "," & DBSet(c_Importe, "N") & "," & DBSet(Mens, "T") & ")"
                        
                        Conn.Execute Sql
                    End If
                End If
            End If '07/02/2011: hasta aqui la parte añadida
        
        End If
    Else
        'MIRAMOS SI EXISTE LA TARJETA
        ' en alzira lo pongo dentro
        codsoc = ""
        '++monica:050508 el numero de tarjeta puede venir a blanco--> dar error
        If Trim(Tarjeta) <> "" Then codsoc = DevuelveDesdeBD("codsocio", "starje", "numtarje", Tarjeta, "N")
        If codsoc = "" Then
            Mens = "Nro. Tarjeta no existe"
            Sql = "insert into tmpinformes (codusu, importe1, fecha1, campo1, campo2, importe2, nombre2, importe3, importe4, importe5, nombre1) values (" & _
                  vSesion.Codigo & "," & DBSet(NumAlbaran, "T") & "," & DBSet(Fecha, "F") & "," & DBSet(Mid(Hora, 1, 2), "N") & _
                  "," & DBSet(Mid(Hora, 3, 2), "N") & "," & DBSet(CodigoCliente, "N") & "," & DBSet(Tarjeta, "T") & "," & DBSet(c_Cantidad, "N") & "," & DBSet(c_Precio, "N") & "," & DBSet(c_Importe, "N") & "," & DBSet(Mens, "T") & ")"
                  
            Conn.Execute Sql
        Else
            ' comprobamos que el socio existe
            ' no haria falta pq hay clave referencial a ssocio
            Sql = ""
            Sql = DevuelveDesdeBDNew(cPTours, "ssocio", "codsocio", "codsocio", codsoc, "N")
            If Sql = "" Then
                Mens = "No existe el cliente"
                Sql = "insert into tmpinformes (codusu, importe1, fecha1, campo1, campo2, importe2, nombre2, importe3, " & _
                      "importe4, importe5, nombre1) values (" & _
                      vSesion.Codigo & "," & DBSet(NumAlbaran, "T") & "," & DBSet(Fecha, "F") & "," & DBSet(Mid(Hora, 1, 2), "N")
                Sql = Sql & "," & DBSet(Mid(Hora, 3, 2), "N") & "," & DBSet(codsoc, "N") & "," & DBSet(codsoc, "T") & "," & _
                        DBSet(c_Cantidad, "N") & "," & DBSet(c_Precio, "N") & "," & DBSet(c_Importe, "N") & "," & DBSet(Mens, "T") & ")"
                
                Conn.Execute Sql
            End If
        End If
    End If
    
    'Comprobamos fechas
    If Not EsFechaOK(Fecha) Then
            Mens = "Fecha incorrecta"
            Sql = "insert into tmpinformes (codusu, importe1, fecha1, campo1, campo2, importe2, nombre2, importe3, " & _
                  "importe4, importe5, nombre1) values (" & _
                  vSesion.Codigo & "," & DBSet(NumAlbaran, "T") & "," & DBSet(Fecha, "F") & "," & DBSet(Mid(Hora, 1, 2), "N")
            Sql = Sql & "," & DBSet(Mid(Hora, 3, 2), "N") & "," & DBSet(CodigoCliente, "N") & "," & DBSet(Fecha, "T") & "," & _
                  DBSet(c_Cantidad, "N") & "," & DBSet(c_Precio, "N") & "," & DBSet(c_Importe, "N") & "," & DBSet(Mens, "T") & ")"
            
            Conn.Execute Sql
    Else
        '[Monica]09/01/2013: en Ribarroja meten todos los turnos del dia a diferencia de Alzira
        If CDate(Fecha) <> CDate(txtCodigo(0).Text) Then
            Mens = "Fecha incorrecta"
            Sql = "insert into tmpinformes (codusu, importe1, fecha1, campo1, campo2, importe2, nombre2, importe3, " & _
                  "importe4, importe5, nombre1) values (" & _
                  vSesion.Codigo & "," & DBSet(NumAlbaran, "T") & "," & DBSet(Fecha, "F") & "," & DBSet(Mid(Hora, 1, 2), "N")
            Sql = Sql & "," & DBSet(Mid(Hora, 3, 2), "N") & "," & DBSet(CodigoCliente, "N") & "," & DBSet(Fecha, "T") & "," & _
                  DBSet(c_Cantidad, "N") & "," & DBSet(c_Precio, "N") & "," & DBSet(c_Importe, "N") & "," & DBSet(Mens, "T") & ")"
            
            Conn.Execute Sql
        End If
    End If
    
    'Comprobamos que el articulo existe en sartic
    Sql = ""
    Sql = DevuelveDesdeBDNew(cPTours, "sartic", "codartic", "codartic", IdProducto, "N")
    If Sql = "" Then
        Mens = "No existe el artículo"
        Dim IdProducto1 As Currency
        IdProducto1 = CCur(IdProducto)
        Sql = "insert into tmpinformes (codusu, importe1, fecha1, campo1, campo2, importe2, nombre2, " & _
              "importe3, importe4, importe5, nombre1) values (" & _
              vSesion.Codigo & "," & DBSet(NumAlbaran, "T") & "," & DBSet(Fecha, "F") & "," & DBSet(Mid(Hora, 1, 2), "N")
        Sql = Sql & "," & DBSet(Mid(Hora, 3, 2), "N") & "," & DBSet(CodigoCliente, "N") & "," & DBSet(IdProducto1, "T") & "," & _
              DBSet(c_Cantidad, "N") & "," & DBSet(c_Precio, "N") & "," & DBSet(c_Importe, "N") & "," & DBSet(Mens, "T") & ")"
              
        Conn.Execute Sql
    End If
    
    'Comprobamos que la forma de pago existe
    If IdTipoPago <> "" Then
        Sql = ""
        Sql = DevuelveDesdeBDNew(cPTours, "sforpa", "codforpa", "forpaalvic", IdTipoPago, "N")
        If Sql = "" Then
            Mens = "No existe la forma de pago Alvic"
            Sql = "insert into tmpinformes (codusu, importe1, fecha1, campo1, campo2, importe2, nombre2, " & _
                  "importe3, importe4, importe5, nombre1) values (" & _
                  vSesion.Codigo & "," & DBSet(NumAlbaran, "T") & "," & DBSet(Fecha, "F") & "," & DBSet(Mid(Hora, 1, 2), "N")
            Sql = Sql & "," & DBSet(Mid(Hora, 3, 2), "N") & "," & DBSet(CodigoCliente, "N") & "," & DBSet(IdTipoPago, "T") & "," & _
                    DBSet(c_Cantidad, "N") & "," & DBSet(c_Precio, "N") & "," & DBSet(c_Importe, "N") & "," & DBSet(Mens, "T") & ")"
            
            Conn.Execute Sql
        End If
    End If
    
    'Comprobamos que el codigo de trabajador existe
    Sql = ""
    Sql = DevuelveDesdeBDNew(cPTours, "straba", "codtraba", "codtraba", IdVendedor, "N")
    If Sql = "" Then
        Mens = "No existe el trabajador"
        Sql = "insert into tmpinformes (codusu, importe1, fecha1, campo1, campo2, importe2, nombre2, " & _
              "importe3, importe4, importe5, nombre1) values (" & _
              vSesion.Codigo & "," & DBSet(NumAlbaran, "T") & "," & DBSet(Fecha, "F") & "," & DBSet(Mid(Hora, 1, 2), "N")
        Sql = Sql & "," & DBSet(Mid(Hora, 3, 2), "N") & "," & DBSet(CodigoCliente, "N") & "," & DBSet(IdVendedor, "T") & "," & _
              DBSet(c_Cantidad, "N") & "," & DBSet(c_Precio, "N") & "," & DBSet(c_Importe, "N") & "," & DBSet(Mens, "T") & ")"
              
        Conn.Execute Sql
    End If
    
    'Comprobamos si hay descuento que el codigo de articulo de dto existe
    If c_Descuento <> 0 Then
        Sql = ""
        Sql = DevuelveDesdeBDNew(cPTours, "sartic", "artdto", "codartic", IdProducto, "N")
        If Sql = "" Then
            Mens = "No tiene artículo de descuento"
            Sql = "insert into tmpinformes (codusu, importe1, fecha1, campo1, campo2, importe2, nombre2, " & _
                  "importe3, importe4, importe5, nombre1) values (" & _
                  vSesion.Codigo & "," & DBSet(NumAlbaran, "T") & "," & DBSet(Fecha, "F") & "," & DBSet(Mid(Hora, 1, 2), "N")
            Sql = Sql & "," & DBSet(Mid(Hora, 3, 2), "N") & "," & DBSet(CodigoCliente, "N") & "," & DBSet(IdProducto, "T") & "," & _
                  DBSet(c_Cantidad, "N") & "," & DBSet(c_Precio, "N") & "," & DBSet(c_Importe, "N") & "," & DBSet(Mens, "T") & ")"
                  
            Conn.Execute Sql
        End If
    End If
    
eComprobarRegistroRib:
    If Err.Number <> 0 Then
        ComprobarRegistroRib = False
    End If
End Function
            
Private Function ComprobarRegistroReg(ByRef RS As Recordset) As Boolean
Dim Sql As String

Dim Base As String
Dim NombreBase As String
Dim Turno As String
Dim NumAlbaran As String
Dim NumFactura As String
Dim IdVendedor As String
Dim NombreVendedor As String
Dim fechahora As String
Dim CodigoCliente As String
Dim NombreCliente As String
Dim Matricula As String
Dim CodigoProducto As String
Dim Surtidor As String
Dim Manguera As String
Dim PrecioLitro As String
Dim PrecioSinDto As String
Dim cantidad As String
Dim Importe As String
Dim IdTipoPago As String
Dim DescrTipoPago As String
Dim CodigoTipoPago As String
Dim NifCliente As String
Dim IdProducto As String
Dim Tarjeta As String
Dim Tarje As String


Dim c_Cantidad As Currency
Dim c_Importe As Currency
Dim c_Precio As Currency

Dim Fecha As String
Dim Hora As String

Dim Mens As String
Dim Kilometros As String


Dim codsoc As String

    On Error GoTo eComprobarRegistro

    ComprobarRegistroReg = True

    Turno = DBLet(RS!Turno, "N")
    
    NumAlbaran = DBLet(RS!Albaran, "N")
    NumFactura = DBLet(RS!Factura, "T")
    If NumFactura <> "" Then
        NumFactura = Mid(NumFactura, 5, Len(NumFactura) - 4)
    End If
    fechahora = DBLet(RS!Fecha, "T")
    Fecha = Mid(fechahora, 7, 2) & "/" & Mid(fechahora, 5, 2) & "/" & Mid(fechahora, 1, 4)
    Hora = Mid(fechahora, 9, 6)
    CodigoCliente = DBLet(RS!CLIENTE, "T")
    NombreCliente = DBLet(RS!nomclien, "T")
    Tarjeta = DBLet(RS!Tarjeta, "N")
    Matricula = DBLet(RS!Matricula, "T")
    IdProducto = DBLet(RS!PRODUCTO, "N")
    Surtidor = DBLet(RS!Surtidor, "N")
    Manguera = DBLet(RS!Manguera, "N")
    
    
    PrecioLitro = DBLet(RS!Precio, "N")
    
    cantidad = DBLet(RS!cantidad, "N")
    Importe = DBLet(RS!Importe, "N")
    IdTipoPago = DBLet(RS!IdTipoPago, "N")
    DescrTipoPago = DBLet(RS!desctipopago, "T")
    CodigoTipoPago = DBLet(RS!IdTipoPago, "N")
    NifCliente = DBLet(RS!NIF, "T")
    
    Kilometros = DBLet(RS!km, "N")
    
    ' en caso de que el codigo de cliente y el nombre no me vengan cojo el asociado a la forma de pago
    If CodigoCliente = "" And NombreCliente = "" Then
        CodigoCliente = DevuelveDesdeBDNew(cPTours, "sforpa", "codsocio", "forpaalvic", IdTipoPago, "N")
        NombreCliente = DevuelveDesdeBDNew(cPTours, "ssocio", "nomsocio", "codsocio", CodigoCliente, "N")
        Tarjeta = CodigoCliente
        If Tarjeta = "0" Then Tarjeta = CodigoCliente
    End If
    '++
    If Mid(CodigoCliente, 1, 2) = "1Z" Then
        CodigoCliente = DevuelveDesdeBDNew(cPTours, "ssocio", "codsocio", "nifsocio", NifCliente, "T")
        If Tarjeta = "0" Then Tarjeta = CodigoCliente
    
    End If
    
    
    If Trim(Importe) = "" Then
        Exit Function
    Else
        If CCur(Importe) = 0 Then Exit Function
    End If
    
    c_Cantidad = cantidad
    c_Importe = Importe
    c_Precio = PrecioLitro
    
    If Trim(NumFactura) <> "" Then
        'SOLAMENTE EN EL CASO DE QUE SEA FACTURA COMPRUEBO QUE EXISTA EL NIF DEL SOCIO
        'Y SI NO EXISTE INTRODUCIRLO EN LA TABLA DE SOCIOS Y TARJETAS
        Tarje = DevuelveDesdeBDNew(cPTours, "ssocio", "codsocio", "nifsocio", NifCliente, "T")
        If Tarje = "" Then
            Tarje = 900000
            Tarje = SugerirCodigoSiguienteStr("ssocio", "codsocio", "codsocio >= 900000 and codsocio <= 999998")
            
'                CtaConta = ""
'                CtaConta = DevuelveDesdeBD("ctaconta", "sparam", "codparam", "01", "N")
            
            Sql = "INSERT INTO ssocio (codsocio, codcoope, nomsocio, domsocio, codposta, pobsocio, " & _
                  "prosocio, nifsocio, telsocio, faxsocio, movsocio, maisocio, wwwsocio, fechaalt, " & _
                  "fechabaj, codtarif, codbanco, codsucur, digcontr, cuentaba, impfactu, dtolitro, " & _
                  "codforpa, tipsocio, bonifbas, bonifesp, codsitua, codmacta, obssocio) VALUES (" & _
                  DBSet(Tarje, "N") & "," & DBSet(vParamAplic.ColecDefecto, "N") & "," & DBSet(NombreCliente, "T") & ",'DESCONOCIDA','46','VALENCIA', " & _
                  "'VALENCIA'," & DBSet(NifCliente, "T") & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & _
                  DBSet(txtCodigo(0).Text, "F") & "," & _
                  ValorNulo & ",0," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & ",0,0," & _
                  "0,0,0,0,0," & DBSet(vParamAplic.CtaContable, "T") & "," & ValorNulo & ")"
                  
            Conn.Execute Sql
                  
            Sql = "INSERT INTO starje (codsocio, numlinea, numtarje, nomtarje, codbanco, codsucur, " & _
                  "digcontr, cuentaba, tiptarje) VALUES (" & DBSet(Tarje, "N") & ",1," & DBSet(Tarje, "N") & "," & DBSet(NombreCliente, "T") & "," & ValorNulo & "," & ValorNulo & "," & _
                  ValorNulo & "," & ValorNulo & ",0)"
            
            Conn.Execute Sql
            
            Tarjeta = Tarje
            
        End If
    End If
    
    'MIRAMOS SI EXISTE LA TARJETA
    If Trim(Tarjeta) <> "0" Then
        codsoc = ""
        codsoc = DevuelveDesdeBD("codsocio", "starje", "numtarje", Tarjeta, "N")
        If codsoc = "" Then
            Mens = "Nro. Tarjeta no existe"
            Sql = "insert into tmpinformes (codusu, importe1, fecha1, campo1, campo2, importe2, nombre2, importe3, importe4, importe5, nombre1) values (" & _
                  vSesion.Codigo & "," & DBSet(NumAlbaran, "T") & "," & DBSet(Fecha, "F") & "," & DBSet(Mid(Hora, 1, 2), "N") & _
                  "," & DBSet(Mid(Hora, 3, 2), "N") & "," & DBSet(CodigoCliente, "N") & "," & DBSet(Tarjeta, "T") & "," & DBSet(c_Cantidad, "N") & "," & DBSet(c_Precio, "N") & "," & DBSet(c_Importe, "N") & "," & DBSet(Mens, "T") & ")"
                  
            Conn.Execute Sql
        End If
    Else
        'COGEMOS LA PRIMERA TARJETA DEPENDIENDO DEL TIPO DE ARTICULO
        Dim tipogaso As String
        tipogaso = DevuelveDesdeBD("tipogaso", "sartic", "codartic", IdProducto, "N")
        Select Case tipogaso
            Case "3" ' bonificado
                Tarje = DevuelveDesdeBDNew(cPTours, "starje", "numtarje", "tiptarje", "1", "N", , "codsocio", CodigoCliente, "N")
                If Tarje = "" Then
                    Mens = "Nro.Tarjeta Bonif.no existe"
                    Sql = "insert into tmpinformes (codusu, importe1, fecha1, campo1, campo2, importe2, nombre2, importe3, importe4, importe5, nombre1) values (" & _
                          vSesion.Codigo & "," & DBSet(NumAlbaran, "T") & "," & DBSet(Fecha, "F") & "," & DBSet(Mid(Hora, 1, 2), "N") & _
                          "," & DBSet(Mid(Hora, 3, 2), "N") & "," & DBSet(CodigoCliente, "N") & "," & DBSet(Tarje, "T") & "," & DBSet(c_Cantidad, "N") & "," & DBSet(c_Precio, "N") & "," & DBSet(c_Importe, "N") & "," & DBSet(Mens, "T") & ")"
                          
                    Conn.Execute Sql
                End If
            Case "0", "1", "2", "4"
                Tarje = DevuelveValor("select numtarje from starje where tiptarje <> 1 and codsocio =" & DBSet(CodigoCliente, "N"))
                If Tarje = "0" Then
                    Mens = "Nro.Tarjeta no existe"
                    Sql = "insert into tmpinformes (codusu, importe1, fecha1, campo1, campo2, importe2, nombre2, importe3, importe4, importe5, nombre1) values (" & _
                          vSesion.Codigo & "," & DBSet(NumAlbaran, "T") & "," & DBSet(Fecha, "F") & "," & DBSet(Mid(Hora, 1, 2), "N") & _
                          "," & DBSet(Mid(Hora, 3, 2), "N") & "," & DBSet(CodigoCliente, "N") & "," & DBSet(Tarje, "T") & "," & DBSet(c_Cantidad, "N") & "," & DBSet(c_Precio, "N") & "," & DBSet(c_Importe, "N") & "," & DBSet(Mens, "T") & ")"
                          
                    Conn.Execute Sql
                End If
        End Select
    End If
    
    'Comprobamos fechas
    If Not EsFechaOK(Fecha) Then
            Mens = "Fecha incorrecta"
            Sql = "insert into tmpinformes (codusu, importe1, fecha1, campo1, campo2, importe2, nombre2, importe3, " & _
                  "importe4, importe5, nombre1) values (" & _
                  vSesion.Codigo & "," & DBSet(NumAlbaran, "T") & "," & DBSet(Fecha, "F") & "," & DBSet(Mid(Hora, 1, 2), "N")
            Sql = Sql & "," & DBSet(Mid(Hora, 3, 2), "N") & "," & DBSet(CodigoCliente, "N") & "," & DBSet(Fecha, "T") & "," & _
                  DBSet(c_Cantidad, "N") & "," & DBSet(c_Precio, "N") & "," & DBSet(c_Importe, "N") & "," & DBSet(Mens, "T") & ")"
            
            Conn.Execute Sql
    Else
        If CDate(Fecha) <> CDate(txtCodigo(0).Text) Then
            Mens = "Fecha incorrecta"
            Sql = "insert into tmpinformes (codusu, importe1, fecha1, campo1, campo2, importe2, nombre2, importe3, " & _
                  "importe4, importe5, nombre1) values (" & _
                  vSesion.Codigo & "," & DBSet(NumAlbaran, "T") & "," & DBSet(Fecha, "F") & "," & DBSet(Mid(Hora, 1, 2), "N")
            Sql = Sql & "," & DBSet(Mid(Hora, 3, 2), "N") & "," & DBSet(CodigoCliente, "N") & "," & DBSet(Fecha, "T") & "," & _
                  DBSet(c_Cantidad, "N") & "," & DBSet(c_Precio, "N") & "," & DBSet(c_Importe, "N") & "," & DBSet(Mens, "T") & ")"
            
            Conn.Execute Sql
        End If
    End If
    
    
    'Comprobamos que el articulo existe en sartic
    Sql = ""
    Sql = DevuelveDesdeBDNew(cPTours, "sartic", "codartic", "codartic", IdProducto, "N")
    If Sql = "" Then
        Mens = "No existe el artículo"
        Sql = "insert into tmpinformes (codusu, importe1, fecha1, campo1, campo2, importe2, nombre2, " & _
              "importe3, importe4, importe5, nombre1) values (" & _
              vSesion.Codigo & "," & DBSet(NumAlbaran, "T") & "," & DBSet(Fecha, "F") & "," & DBSet(Mid(Hora, 1, 2), "N")
        Sql = Sql & "," & DBSet(Mid(Hora, 3, 2), "N") & "," & DBSet(CodigoCliente, "N") & "," & DBSet(IdProducto, "T") & "," & _
              DBSet(c_Cantidad, "N") & "," & DBSet(c_Precio, "N") & "," & DBSet(c_Importe, "N") & "," & DBSet(Mens, "T") & ")"
              
        Conn.Execute Sql
    End If
    
    
    'Comprobamos que el socio existe
    If CodigoCliente <> "" Then
        Sql = ""
        Sql = DevuelveDesdeBDNew(cPTours, "ssocio", "codsocio", "codsocio", CodigoCliente, "N")
        If Sql = "" Then
            Mens = "No existe el cliente"
            Sql = "insert into tmpinformes (codusu, importe1, fecha1, campo1, campo2, importe2, nombre2, importe3, " & _
                  "importe4, importe5, nombre1) values (" & _
                  vSesion.Codigo & "," & DBSet(NumAlbaran, "T") & "," & DBSet(Fecha, "F") & "," & DBSet(Mid(Hora, 1, 2), "N")
            Sql = Sql & "," & DBSet(Mid(Hora, 3, 2), "N") & "," & DBSet(CodigoCliente, "N") & "," & DBSet(CodigoCliente, "T") & "," & _
                    DBSet(c_Cantidad, "N") & "," & DBSet(c_Precio, "N") & "," & DBSet(c_Importe, "N") & "," & DBSet(Mens, "T") & ")"
            
            Conn.Execute Sql
        End If
    End If
    
    'Comprobamos que la forma de pago existe
    If IdTipoPago <> "" Then
        Sql = ""
        Sql = DevuelveDesdeBDNew(cPTours, "sforpa", "codforpa", "forpaalvic", IdTipoPago, "N")
        
        
        If Sql = "" Then
            
            '[Monica]05/01/2015: si el socio es de catadau o llombai cogemos su forma de pago (la del cliente)
            Sql = "select codforpa from ssocio where codsocio = " & DBSet(CodigoCliente, "N") & " and codcoope in (1,2) "
            If TotalRegistrosConsulta(Sql) <> 0 Then Exit Function
            
            
            Mens = "No existe la forma de pago Alvic"
            Sql = "insert into tmpinformes (codusu, importe1, fecha1, campo1, campo2, importe2, nombre2, " & _
                  "importe3, importe4, importe5, nombre1) values (" & _
                  vSesion.Codigo & "," & DBSet(NumAlbaran, "T") & "," & DBSet(Fecha, "F") & "," & DBSet(Mid(Hora, 1, 2), "N")
            Sql = Sql & "," & DBSet(Mid(Hora, 3, 2), "N") & "," & DBSet(CodigoCliente, "N") & "," & DBSet(IdTipoPago, "T") & "," & _
                    DBSet(c_Cantidad, "N") & "," & DBSet(c_Precio, "N") & "," & DBSet(c_Importe, "N") & "," & DBSet(Mens, "T") & ")"
            
            Conn.Execute Sql
        End If
    End If
    
eComprobarRegistro:
    If Err.Number <> 0 Then
        ComprobarRegistroReg = False
    End If
End Function
            
            
            
            
            
Private Function InsertarLinea(cad As String) As Boolean
Dim NumLin As String
Dim codpro As String
Dim articulo As String
Dim Familia As String
Dim Precio As String
Dim ImpDes As String
Dim CodIVA As String
Dim b As Boolean
Dim Codclave As String
Dim Sql As String

Dim Import As Currency

Dim Base As String
Dim NombreBase As String
Dim Turno As String
Dim NumAlbaran As String
Dim NumFactura As String
Dim IdVendedor As String
Dim NombreVendedor As String
Dim fechahora As String
Dim Fecha As String
Dim Hora As String
Dim CodigoCliente As String
Dim NombreCliente As String
Dim Matricula As String
Dim Tarjeta As String
Dim CodigoProducto As String
Dim Surtidor As String
Dim Manguera As String
Dim PrecioLitro As String
Dim Descuento As String
Dim cantidad As String
Dim Importe As String
Dim IdTipoPago As String
Dim DescrTipoPago As String
Dim CodigoTipoPago As String
Dim NifCliente As String
Dim IdProducto As String

Dim c_Cantidad As Currency
Dim c_Importe As Currency
Dim c_Precio As Currency
Dim c_Descuento As Currency

Dim Tarje As String


Dim Mens As String
Dim numlinea As Long

Dim codsoc As String
Dim forpa As String

Dim Kilometros As String
Dim NomArtic As String

    On Error GoTo eInsertarLinea

    InsertarLinea = True
    

    Base = Mid(cad, 1, 10)
    NombreBase = Mid(cad, 11, 50)
    Turno = Mid(cad, 982, 10) 'txtcodigo(1).Text 'el turno que yo le diga, antes: Mid(cad, 61, 10)
    If CByte(Turno) > 9 Then Turno = "9"
    NumAlbaran = Mid(cad, 72, 19)
    NumFactura = Mid(cad, 94, 17)
    IdVendedor = Mid(cad, 121, 10)
    NombreVendedor = Mid(cad, 131, 50)
    fechahora = Mid(cad, 181, 14)
    Fecha = Mid(fechahora, 7, 2) & "/" & Mid(fechahora, 5, 2) & "/" & Mid(fechahora, 1, 4)
    Hora = Mid(fechahora, 9, 2) & ":" & Mid(fechahora, 11, 2) & ":" & Mid(fechahora, 13, 2)
    CodigoCliente = Mid(cad, 195, 20)
    NombreCliente = Mid(cad, 215, 70)
    Tarjeta = Mid(cad, 290, 20)
    Matricula = Mid(cad, 370, 20)
    IdProducto = Mid(cad, 493, 20)
    Surtidor = Mid(cad, 538, 10)
    Manguera = Mid(cad, 548, 10)
    '[Monica]24/08/2015: ahora el precio es el de la posicion 864 antes era sin el de la 568
    PrecioLitro = Mid(cad, 864, 18)
    '[Monica]24/08/2015: añadimos el descuento
    Descuento = Mid(cad, 586, 18)
    
    cantidad = Mid(cad, 650, 18)
    Importe = Mid(cad, 668, 18)
    IdTipoPago = Mid(cad, 784, 10)
    DescrTipoPago = Mid(cad, 794, 25)
    CodigoTipoPago = Mid(cad, 1, 10)
    NifCliente = Mid(cad, 834, 9)
    
    '[Monica]24/06/2013: introducimos los kms em el traspaso
    Kilometros = Mid(cad, 415, 18)
    
    
    
    If Trim(Importe) = "" Then
        Exit Function
    Else
        If CCur(Importe) = 0 Then Exit Function
    End If
    
    c_Cantidad = Round2(CCur(cantidad) / 100, 2)
    c_Importe = Round2(CCur(Importe) / 100, 2)
    c_Precio = Round2(CCur(PrecioLitro) / 100000, 5)
    c_Descuento = Round2(CCur(Descuento) / 100000, 5)
    
'    '### [Monica] 17/09/2007
'    'no insertamos aquellas lineas de albaran de importe = 0
'    Importe = DBSet(c_Importe, "N")
'    If Import = 0 Then
'        InsertarLinea = True
'        Exit Function
'    End If
'    'hasta aqui
    
    'VRS:4.0.1(0) actualizamos el precio de articulo
    Sql = "update sartic set preventa = " & DBSet(c_Precio, "N") & _
          " where codartic = " & DBSet(IdProducto, "N")
    Conn.Execute Sql
    
    If DevuelveValor("select ctrstock from sartic where codartic = " & DBSet(IdProducto, "N")) = 1 Then
        Sql = "update sartic set " & _
              "  canstock = canstock - " & DBSet(c_Cantidad, "N") & _
              " where codartic = " & DBSet(IdProducto, "N")
        Conn.Execute Sql
    End If
    
    ' insertamos en la tabla de albaranes
    Codclave = SugerirCodigoSiguienteStr("scaalb", "codclave")
    
    forpa = ""
    forpa = DevuelveDesdeBDNew(cPTours, "sforpa", "codforpa", "forpaalvic", IdTipoPago, "N")
    
    If Trim(NumFactura) <> "" Then
        codsoc = DevuelveDesdeBDNew(cPTours, "ssocio", "codsocio", "nifsocio", NifCliente, "T")
        '[Monica]17/06/2013: miramos si la tarjeta viene con algun asterisco
        If Mid(Tarjeta, 1, 4) = "****" Or Trim(Tarjeta) = "" Or InStr(1, Tarjeta, "*") <> 0 Then
            Tarjeta = codsoc
        Else '++monica: 15/02/2008 las tarjetas profesionales tienen 16 caracteres solo analizo los 8 últimos
            If Len(Trim(Tarjeta)) = 16 Then
                Tarjeta = Mid(Tarjeta, 9, 16)
            End If
            '++
        End If
        'fechahora--> txtcodigo(0).Text & " " & Time
        
        Sql = "INSERT INTO scaalb (codclave, codsocio, numtarje, numalbar, fecalbar, horalbar, " & _
              "codturno, codartic, cantidad, preciove, importel, codforpa, matricul, codtraba, " & _
              "numfactu, numlinea, kilometros, dtoalvic) VALUES (" & DBSet(Codclave, "T") & "," & DBSet(codsoc, "N") & "," & DBSet(Tarjeta, "N") & "," & _
               DBSet(NumAlbaran, "T") & "," & DBSet(txtCodigo(0).Text, "F") & "," & DBSet(txtCodigo(0).Text & " " & Hora, "FH") & "," & DBSet(txtCodigo(1).Text, "N") & "," & _
               DBSet(IdProducto, "N") & "," & DBSet(c_Cantidad, "N") & "," & DBSet(c_Precio, "N") & "," & _
               DBSet(c_Importe, "N") & "," & DBSet(forpa, "N") & "," & DBSet(Matricula, "T") & "," & DBSet(IdVendedor, "N") & ","
    
        numlinea = SugerirCodigoSiguienteStr("scaalb", "numlinea", "numfactu = " & DBSet(NumFactura, "N"))
        Sql = Sql & DBSet(NumFactura, "N") & "," & DBSet(numlinea, "N") & ","
    Else
        '[Monica]16/01/2014: si me viene una factura tpv sin nro pregunto sobre que cliente la pongo visa o contado
        If InStr(1, CodigoCliente, "1Z") <> 0 Then
            NomArtic = DevuelveDesdeBDNew(cPTours, "sartic", "nomartic", "codartic", IdProducto, "N")
            If MsgBox("Factura de cliente de paso sin número de factura. " & vbCrLf & vbCrLf & "Albaran: " & NumAlbaran & vbCrLf & "Fecha: " & txtCodigo(0).Text & " " & Hora & vbCrLf & "Articulo: " & NomArtic & vbCrLf & "Importe: " & c_Importe & vbCrLf & vbCrLf & "¿ Asignar a ventas contado ? " & vbCrLf & "(en caso negativo se asignará a ventas tarjeta)" & vbCrLf & vbCrLf, vbQuestion + vbYesNo + vbDefaultButton1) = vbYes Then
                Tarjeta = "900000"
            Else
                Tarjeta = "900002"
            End If
            
            Sql = "INSERT INTO scaalb (codclave, codsocio, numtarje, numalbar, fecalbar, horalbar, " & _
                  "codturno, codartic, cantidad, preciove, importel, codforpa, matricul, codtraba, " & _
                  "numfactu, numlinea, kilometros, dtoalvic) VALUES (" & DBSet(Codclave, "T") & "," & DBSet(Tarjeta, "N") & "," & DBSet(Tarjeta, "N") & "," & _
                   DBSet(NumAlbaran, "T") & "," & DBSet(txtCodigo(0).Text, "F") & "," & DBSet(txtCodigo(0).Text & " " & Hora, "FH") & "," & DBSet(txtCodigo(1).Text, "N") & "," & _
                   DBSet(IdProducto, "N") & "," & DBSet(c_Cantidad, "N") & "," & DBSet(c_Precio, "N") & "," & _
                   DBSet(c_Importe, "N") & "," & DBSet(forpa, "N") & "," & DBSet(Matricula, "T") & "," & DBSet(IdVendedor, "N") & ","
            Sql = Sql & "0,0,"
        Else
        
            '[Monica]17/06/2013: miramos si la tarjeta viene con algun asterisco
            If Mid(Tarjeta, 1, 4) = "****" Or Trim(Tarjeta) = "" Or InStr(1, Tarjeta, "*") <> 0 Then
                Tarjeta = CodigoCliente
            Else '++monica: 15/02/2008 las tarjetas profesionales tienen 16 caracteres solo analizo los 8 últimos
                If Len(Trim(Tarjeta)) = 16 Then
                    Tarjeta = Mid(Tarjeta, 9, 16)
                End If
                '++
            End If
            Sql = "INSERT INTO scaalb (codclave, codsocio, numtarje, numalbar, fecalbar, horalbar, " & _
                  "codturno, codartic, cantidad, preciove, importel, codforpa, matricul, codtraba, " & _
                  "numfactu, numlinea, kilometros, dtoalvic) VALUES (" & DBSet(Codclave, "T") & "," & DBSet(CodigoCliente, "N") & "," & DBSet(Tarjeta, "N") & "," & _
                   DBSet(NumAlbaran, "T") & "," & DBSet(txtCodigo(0).Text, "F") & "," & DBSet(txtCodigo(0).Text & " " & Hora, "FH") & "," & DBSet(txtCodigo(1).Text, "N") & "," & _
                   DBSet(IdProducto, "N") & "," & DBSet(c_Cantidad, "N") & "," & DBSet(c_Precio, "N") & "," & _
                   DBSet(c_Importe, "N") & "," & DBSet(forpa, "N") & "," & DBSet(Matricula, "T") & "," & DBSet(IdVendedor, "N") & ","
            Sql = Sql & "0,0,"
            
        End If
    End If
    
    '[monica]24/06/2013: añadimos los kilometros
    Sql = Sql & DBSet(Round2(ComprobarCero(Trim(Kilometros)) / 100, 0), "N", "S") & "," '& ")"
 
 
    '[Monica]24/08/2015: añadimos el descuento
    Sql = Sql & DBSet(c_Descuento, "N") & ")"
 
    Conn.Execute Sql
    
eInsertarLinea:
    If Err.Number <> 0 Then
        InsertarLinea = False
        MsgBox "Error en Insertar Linea " & Err.Description, vbExclamation
    End If
End Function
            
            
Private Function InsertarLineaAlz(cad As String) As Boolean
Dim NumLin As String
Dim codpro As String
Dim articulo As String
Dim Familia As String
Dim Precio As String
Dim ImpDes As String
Dim CodIVA As String
Dim b As Boolean
Dim Codclave As String
Dim Sql As String

Dim Import As Currency

Dim Base As String
Dim NombreBase As String
Dim Turno As String
Dim NumAlbaran As String
Dim NumFactura As String
Dim IdVendedor As String
Dim NombreVendedor As String
Dim fechahora As String
Dim Fecha As String
Dim Hora As String
Dim CodigoCliente As String
Dim NombreCliente As String
Dim Matricula As String
Dim Tarjeta As String
Dim CodigoProducto As String
Dim Surtidor As String
Dim Manguera As String
Dim PrecioLitro As String
Dim cantidad As String
Dim Importe As String
Dim Descuento As String
Dim IdTipoPago As String
Dim DescrTipoPago As String
Dim CodigoTipoPago As String
Dim NifCliente As String
Dim IdProducto As String

Dim c_Cantidad As Currency
Dim c_Importe As Currency
Dim c_Importe1 As Currency
Dim c_Importe2 As Currency
Dim c_Precio As Currency
Dim c_Precio2 As Currency
Dim c_Descuento As Currency
Dim IdProductoDes As String

Dim Tarje As String


Dim Mens As String
Dim numlinea As Long

Dim codsoc As String
Dim forpa As String
Dim Kilometros As String


    On Error GoTo eInsertarLineaAlz

    InsertarLineaAlz = True
    

    Base = Mid(cad, 1, 10)
    NombreBase = Mid(cad, 11, 50)
    Turno = Mid(cad, 982, 10) 'txtcodigo(1).Text 'el turno que yo le diga, antes: Mid(cad, 61, 10)
    If CByte(Turno) > 9 Then Turno = "9"
    NumAlbaran = Mid(cad, 71, 20)
    NumFactura = Mid(cad, 94, 17)
    IdVendedor = Mid(cad, 121, 10)
    NombreVendedor = Mid(cad, 131, 50)
    fechahora = Mid(cad, 181, 14)
    Fecha = Mid(fechahora, 7, 2) & "/" & Mid(fechahora, 5, 2) & "/" & Mid(fechahora, 1, 4)
    Hora = Mid(fechahora, 9, 2) & ":" & Mid(fechahora, 11, 2) & ":" & Mid(fechahora, 13, 2)
'    CodigoCliente = Mid(cad, 195, 20)
    NombreCliente = Mid(cad, 215, 70)
'    Tarjeta = Mid(cad, 290, 20)
    Tarjeta = Mid(cad, 195, 20)
    Matricula = Mid(cad, 370, 20)
    IdProducto = Mid(cad, 493, 20)
    Surtidor = Mid(cad, 538, 10)
    Manguera = Mid(cad, 548, 10)
    PrecioLitro = Mid(cad, 568, 18)
    cantidad = Mid(cad, 650, 18)
    Importe = Mid(cad, 668, 18)
    Descuento = Mid(cad, 586, 18)
    IdTipoPago = Mid(cad, 784, 10)
    DescrTipoPago = Mid(cad, 794, 25)
    CodigoTipoPago = Mid(cad, 1, 10)
    NifCliente = Mid(cad, 834, 9)
    
    '[Monica]24/06/2013: introducimos los kms em el traspaso
    Kilometros = Mid(cad, 415, 18)
    
    If Trim(Importe) = "" Then
        Exit Function
    Else
        If CCur(Importe) = 0 Then Exit Function
    End If
    
    c_Cantidad = Round2(CCur(cantidad) / 100, 2)
    c_Importe = Round2(CCur(Importe) / 100, 2)
    c_Precio = Round2(CCur(PrecioLitro) / 100000, 5)

    If Trim(Descuento) <> "" Then
        If CCur(Descuento) <> 0 Then
            c_Descuento = Round2(CCur(Descuento) / 100000, 5)
            c_Importe1 = Round2(c_Cantidad * c_Precio, 2)
            c_Importe2 = c_Importe - c_Importe1
            c_Importe = c_Importe1
            c_Precio2 = Round2(c_Importe2 / c_Cantidad * (-1), 3)
            IdProductoDes = DevuelveDesdeBDNew(cPTours, "sartic", "artdto", "codartic", IdProducto, "N")
        Else
            c_Descuento = 0
        End If
    End If

    
'    '### [Monica] 17/09/2007
'    'no insertamos aquellas lineas de albaran de importe = 0
'    Importe = DBSet(c_Importe, "N")
'    If Import = 0 Then
'        InsertarLineaAlz = True
'        Exit Function
'    End If
'    'hasta aqui
    
    'VRS:4.0.1(0) actualizamos el precio de articulo
    Sql = "update sartic set preventa = " & DBSet(c_Precio, "N") & _
          ", canstock = canstock - " & DBSet(c_Cantidad, "N") & _
          " where codartic = " & DBSet(IdProducto, "N")
    Conn.Execute Sql
    
'    If DevuelveValor("select ctrstock from sartic where codartic = " & DBSet(IdProducto, "N")) = 1 Then
'        SQL = "update sartic set " & _
'              "  canstock = canstock - " & DBSet(c_Cantidad, "N") & _
'              " where codartic = " & DBSet(IdProducto, "N")
'        Conn.Execute SQL
'    End If
    
    
    ' insertamos en la tabla de albaranes
    Codclave = SugerirCodigoSiguienteStr("scaalb", "codclave")
    
    forpa = ""
    forpa = DevuelveDesdeBDNew(cPTours, "sforpa", "codforpa", "forpaalvic", IdTipoPago, "N")
    
    '[Monica]30/11/2011 añadida segunda condicion
    If Trim(NumFactura) <> "" And InStr(1, Tarjeta, "Z") <> 0 Then
        codsoc = DevuelveDesdeBDNew(cPTours, "ssocio", "codsocio", "nifsocio", NifCliente, "T")
        If Mid(Tarjeta, 1, 4) = "****" Or Trim(Tarjeta) = "" Then
            Tarjeta = codsoc
            
        Else '[Monica]07/02/2011 buscamos la tarjeta que corresponda para meter pq me viene Z
            If codsoc >= 900000 Then
                Tarjeta = DevuelveValor("select numtarje from starje where codsocio= " & DBSet(codsoc, "N"))
            Else
                ' el socio es inferior a 900000 miro si hay tarjeta dependiendo del producto
                Dim TipArtic As Integer
                TipArtic = DevuelveValor("select tipogaso from sartic where codartic = " & DBSet(IdProducto, "N"))
                If TipArtic = 3 Then ' si el articulo es gasoleo bonificado
                    Tarjeta = DevuelveValor("select numtarje from starje where codsocio= " & DBSet(codsoc, "N") & " and tiptarje = 1")
                    
                    '[Monica]22/11/2011: si no tiene tarjeta de gasoleo bonificado le meto la primera tarjeta que encuentre
                    If Tarjeta = "0" Then
                        Tarjeta = DevuelveValor("select numtarje from starje where codsocio = " & DBSet(codsoc, "N"))
                    End If
                    
                Else
                    Tarjeta = DevuelveValor("select numtarje from starje where codsocio= " & DBSet(codsoc, "N") & " and tiptarje = 0")
                End If
            End If
            '[Monica]07/02/2011 hasta aqui
        End If
        'fechahora--> txtcodigo(0).Text & " " & Time
        
        Sql = "INSERT INTO scaalb (codclave, codsocio, numtarje, numalbar, fecalbar, horalbar, " & _
              "codturno, codartic, cantidad, preciove, importel, codforpa, matricul, codtraba, " & _
              "numfactu, numlinea, kilometros) VALUES (" & DBSet(Codclave, "T") & "," & DBSet(codsoc, "N") & "," & DBSet(Tarjeta, "N") & "," & _
               DBSet(NumAlbaran, "T") & "," & DBSet(txtCodigo(0).Text, "F") & "," & DBSet(txtCodigo(0).Text & " " & Hora, "FH") & "," & DBSet(txtCodigo(1).Text, "N") & "," & _
               DBSet(IdProducto, "N") & "," & DBSet(c_Cantidad, "N") & "," & DBSet(c_Precio, "N") & "," & _
               DBSet(c_Importe, "N") & "," & DBSet(forpa, "N") & "," & DBSet(Matricula, "T") & "," & DBSet(IdVendedor, "N") & ","
    
        numlinea = SugerirCodigoSiguienteStr("scaalb", "numlinea", "numfactu = " & DBSet(NumFactura, "N"))
        Sql = Sql & DBSet(NumFactura, "N") & "," & DBSet(numlinea, "N") & ","
        
        '[monica]24/06/2013: añadimos los kilometros
        Sql = Sql & DBSet(Round2(ComprobarCero(Trim(Kilometros)) / 100, 0), "N", "S") & ")"
   
        Conn.Execute Sql
        
        If c_Descuento <> 0 Then
            Sql = "update sartic set preventa = " & DBSet(c_Precio2, "N") & _
                  " where codartic = " & DBSet(IdProductoDes, "N")
            Conn.Execute Sql
            
            Codclave = SugerirCodigoSiguienteStr("scaalb", "codclave")
           
            Sql = "INSERT INTO scaalb (codclave, codsocio, numtarje, numalbar, fecalbar, horalbar, " & _
                  "codturno, codartic, cantidad, preciove, importel, codforpa, matricul, codtraba, " & _
                  "numfactu, numlinea, kilometros) VALUES (" & DBSet(Codclave, "T") & "," & DBSet(codsoc, "N") & "," & DBSet(codsoc, "N") & "," & _
                   DBSet(NumAlbaran, "T") & "," & DBSet(txtCodigo(0).Text, "F") & "," & DBSet(txtCodigo(0).Text & " " & Hora, "FH") & "," & DBSet(txtCodigo(1).Text, "N") & "," & _
                   DBSet(IdProductoDes, "N") & "," & DBSet(c_Cantidad, "N") & "," & DBSet(c_Precio2, "N") & "," & _
                   DBSet(c_Importe2, "N") & "," & DBSet(forpa, "N") & "," & DBSet(Matricula, "T") & "," & DBSet(IdVendedor, "N") & ","
        
            numlinea = numlinea + 1
            Sql = Sql & DBSet(NumFactura, "N") & "," & DBSet(numlinea, "N") & ","
            
            '[monica]24/06/2013: añadimos los kilometros
            Sql = Sql & DBSet(Round2(ComprobarCero(Trim(Kilometros)) / 100, 0), "N", "S") & ")"
        
            Conn.Execute Sql
        End If
        
    Else
        '[Monica]30/11/2010
        If Trim(NumFactura) <> "" Then
            codsoc = DevuelveDesdeBDNew(cPTours, "starje", "codsocio", "numtarje", Tarjeta, "N")
            If Mid(Tarjeta, 1, 4) = "****" Or Trim(Tarjeta) = "" Then
                Tarjeta = codsoc
            End If
            'fechahora--> txtcodigo(0).Text & " " & Time
            
            Sql = "INSERT INTO scaalb (codclave, codsocio, numtarje, numalbar, fecalbar, horalbar, " & _
                  "codturno, codartic, cantidad, preciove, importel, codforpa, matricul, codtraba, " & _
                  "numfactu, numlinea, kilometros) VALUES (" & DBSet(Codclave, "T") & "," & DBSet(codsoc, "N") & "," & DBSet(Tarjeta, "N") & "," & _
                   DBSet(NumAlbaran, "T") & "," & DBSet(txtCodigo(0).Text, "F") & "," & DBSet(txtCodigo(0).Text & " " & Hora, "FH") & "," & DBSet(txtCodigo(1).Text, "N") & "," & _
                   DBSet(IdProducto, "N") & "," & DBSet(c_Cantidad, "N") & "," & DBSet(c_Precio, "N") & "," & _
                   DBSet(c_Importe, "N") & "," & DBSet(forpa, "N") & "," & DBSet(Matricula, "T") & "," & DBSet(IdVendedor, "N") & ","
        
            numlinea = SugerirCodigoSiguienteStr("scaalb", "numlinea", "numfactu = " & DBSet(NumFactura, "N"))
            Sql = Sql & DBSet(NumFactura, "N") & "," & DBSet(numlinea, "N") & ","
            
            '[monica]24/06/2013: añadimos los kilometros
            Sql = Sql & DBSet(Round2(ComprobarCero(Trim(Kilometros)) / 100, 0), "N", "S") & ")"
            
            
            Conn.Execute Sql
            
            If c_Descuento <> 0 Then
                Sql = "update sartic set preventa = " & DBSet(c_Precio2, "N") & _
                      " where codartic = " & DBSet(IdProductoDes, "N")
                Conn.Execute Sql
                
                Codclave = SugerirCodigoSiguienteStr("scaalb", "codclave")
               
                Sql = "INSERT INTO scaalb (codclave, codsocio, numtarje, numalbar, fecalbar, horalbar, " & _
                      "codturno, codartic, cantidad, preciove, importel, codforpa, matricul, codtraba, " & _
                      "numfactu, numlinea) VALUES (" & DBSet(Codclave, "T") & "," & DBSet(codsoc, "N") & "," & DBSet(Tarjeta, "N") & "," & _
                       DBSet(NumAlbaran, "T") & "," & DBSet(txtCodigo(0).Text, "F") & "," & DBSet(txtCodigo(0).Text & " " & Hora, "FH") & "," & DBSet(txtCodigo(1).Text, "N") & "," & _
                       DBSet(IdProductoDes, "N") & "," & DBSet(c_Cantidad, "N") & "," & DBSet(c_Precio2, "N") & "," & _
                       DBSet(c_Importe2, "N") & "," & DBSet(forpa, "N") & "," & DBSet(Matricula, "T") & "," & DBSet(IdVendedor, "N") & ","
            
                numlinea = numlinea + 1
                Sql = Sql & DBSet(NumFactura, "N") & "," & DBSet(numlinea, "N") & ")"
            
                Conn.Execute Sql
            End If
        
        Else
            CodigoCliente = DevuelveDesdeBDNew(cPTours, "starje", "codsocio", "numtarje", Tarjeta, "N")
            Sql = "INSERT INTO scaalb (codclave, codsocio, numtarje, numalbar, fecalbar, horalbar, " & _
                  "codturno, codartic, cantidad, preciove, importel, codforpa, matricul, codtraba, " & _
                  "numfactu, numlinea, kilometros) VALUES (" & DBSet(Codclave, "T") & "," & DBSet(CodigoCliente, "N") & "," & DBSet(Tarjeta, "N") & "," & _
                   DBSet(NumAlbaran, "T") & "," & DBSet(txtCodigo(0).Text, "F") & "," & DBSet(txtCodigo(0).Text & " " & Hora, "FH") & "," & DBSet(txtCodigo(1).Text, "N") & "," & _
                   DBSet(IdProducto, "N") & "," & DBSet(c_Cantidad, "N") & "," & DBSet(c_Precio, "N") & "," & _
                   DBSet(c_Importe, "N") & "," & DBSet(forpa, "N") & "," & DBSet(Matricula, "T") & "," & DBSet(IdVendedor, "N") & ","
            Sql = Sql & "0,0,"
            
            '[monica]24/06/2013: añadimos los kilometros
            Sql = Sql & DBSet(Round2(ComprobarCero(Trim(Kilometros)) / 100, 0), "N", "S") & ")"
            
            
            Conn.Execute Sql
            
            If c_Descuento <> 0 Then
                Sql = "update sartic set preventa = " & DBSet(c_Precio2, "N") & _
                      " where codartic = " & DBSet(IdProductoDes, "N")
                Conn.Execute Sql
                
                Codclave = SugerirCodigoSiguienteStr("scaalb", "codclave")
                
                Sql = "INSERT INTO scaalb (codclave, codsocio, numtarje, numalbar, fecalbar, horalbar, " & _
                      "codturno, codartic, cantidad, preciove, importel, codforpa, matricul, codtraba, " & _
                      "numfactu, numlinea) VALUES (" & DBSet(Codclave, "T") & "," & DBSet(CodigoCliente, "N") & "," & DBSet(Tarjeta, "N") & "," & _
                       DBSet(NumAlbaran, "T") & "," & DBSet(txtCodigo(0).Text, "F") & "," & DBSet(txtCodigo(0).Text & " " & Hora, "FH") & "," & DBSet(txtCodigo(1).Text, "N") & "," & _
                       DBSet(IdProductoDes, "N") & "," & DBSet(c_Cantidad, "N") & "," & DBSet(c_Precio2, "N") & "," & _
                       DBSet(c_Importe2, "N") & "," & DBSet(forpa, "N") & "," & DBSet(Matricula, "T") & "," & DBSet(IdVendedor, "N") & ","
                Sql = Sql & "0,0)"
            
                Conn.Execute Sql
            End If
        End If
    End If
 
    
    
eInsertarLineaAlz:
    If Err.Number <> 0 Then
        InsertarLineaAlz = False
        MsgBox "Error en Insertar Linea " & Err.Description, vbExclamation
    End If
End Function
            
' la diferencia con Alzira es que inserta el turno del fichero no del frame
Private Function InsertarLineaRib(cad As String) As Boolean
Dim NumLin As String
Dim codpro As String
Dim articulo As String
Dim Familia As String
Dim Precio As String
Dim ImpDes As String
Dim CodIVA As String
Dim b As Boolean
Dim Codclave As String
Dim Sql As String

Dim Import As Currency

Dim Base As String
Dim NombreBase As String
Dim Turno As String
Dim NumAlbaran As String
Dim NumFactura As String
Dim IdVendedor As String
Dim NombreVendedor As String
Dim fechahora As String
Dim Fecha As String
Dim Hora As String
Dim CodigoCliente As String
Dim NombreCliente As String
Dim Matricula As String
Dim Tarjeta As String
Dim CodigoProducto As String
Dim Surtidor As String
Dim Manguera As String
Dim PrecioLitro As String
Dim cantidad As String
Dim Importe As String
Dim Descuento As String
Dim IdTipoPago As String
Dim DescrTipoPago As String
Dim CodigoTipoPago As String
Dim NifCliente As String
Dim IdProducto As String

Dim c_Cantidad As Currency
Dim c_Importe As Currency
Dim c_Importe1 As Currency
Dim c_Importe2 As Currency
Dim c_Precio As Currency
Dim c_Precio2 As Currency
Dim c_Descuento As Currency
Dim IdProductoDes As String

Dim Tarje As String


Dim Mens As String
Dim numlinea As Long

Dim codsoc As String
Dim forpa As String

Dim Kilometros As String

    On Error GoTo eInsertarLineaRib

    InsertarLineaRib = True
    

    Base = Mid(cad, 1, 10)
    NombreBase = Mid(cad, 11, 50)
    Turno = Mid(cad, 982, 10) 'txtcodigo(1).Text 'el turno que yo le diga, antes: Mid(cad, 61, 10)
    If CByte(Turno) > 9 Then Turno = "9"
    NumAlbaran = Mid(cad, 71, 20)
    NumFactura = Mid(cad, 92, 7) '14/05/2013 antes 94,17
    IdVendedor = Mid(cad, 121, 10)
    NombreVendedor = Mid(cad, 131, 50)
    fechahora = Mid(cad, 181, 14)
    Fecha = Mid(fechahora, 7, 2) & "/" & Mid(fechahora, 5, 2) & "/" & Mid(fechahora, 1, 4)
    Hora = Mid(fechahora, 9, 2) & ":" & Mid(fechahora, 11, 2) & ":" & Mid(fechahora, 13, 2)
'    CodigoCliente = Mid(cad, 195, 20)
    NombreCliente = Mid(cad, 215, 70)
'    Tarjeta = Mid(cad, 290, 20)
    Tarjeta = Mid(cad, 195, 20)
    Matricula = Mid(cad, 370, 20)
    IdProducto = Mid(cad, 493, 20)
    Surtidor = Mid(cad, 538, 10)
    Manguera = Mid(cad, 548, 10)
    PrecioLitro = Mid(cad, 568, 18)
    cantidad = Mid(cad, 650, 18)
    Importe = Mid(cad, 668, 18)
    Descuento = Mid(cad, 586, 18)
    IdTipoPago = Mid(cad, 784, 10)
    DescrTipoPago = Mid(cad, 794, 25)
    CodigoTipoPago = Mid(cad, 1, 10)
    NifCliente = Mid(cad, 834, 9)
    
    '[Monica]24/06/2013: insertamos los kilometros
    Kilometros = Mid(cad, 415, 18)
    
    
    If Trim(Importe) = "" Then
        Exit Function
    Else
        If CCur(Importe) = 0 Then Exit Function
    End If
    
    c_Cantidad = Round2(CCur(cantidad) / 100, 2)
    c_Importe = Round2(CCur(Importe) / 100, 2)
    c_Precio = Round2(CCur(PrecioLitro) / 100000, 5)

    If Trim(Descuento) <> "" Then
        If CCur(Descuento) <> 0 Then
            c_Descuento = Round2(CCur(Descuento) / 100000, 5)
            c_Importe1 = Round2(c_Cantidad * c_Precio, 2)
            c_Importe2 = c_Importe - c_Importe1
            c_Importe = c_Importe1
            c_Precio2 = Round2(c_Importe2 / c_Cantidad * (-1), 3)
            IdProductoDes = DevuelveDesdeBDNew(cPTours, "sartic", "artdto", "codartic", IdProducto, "N")
        Else
            c_Descuento = 0
        End If
    End If

    
'    '### [Monica] 17/09/2007
'    'no insertamos aquellas lineas de albaran de importe = 0
'    Importe = DBSet(c_Importe, "N")
'    If Import = 0 Then
'        InsertarLineaAlz = True
'        Exit Function
'    End If
'    'hasta aqui
    
    'VRS:4.0.1(0) actualizamos el precio de articulo
    Sql = "update sartic set preventa = " & DBSet(c_Precio, "N") & _
          ", canstock = canstock - " & DBSet(c_Cantidad, "N") & _
          " where codartic = " & DBSet(IdProducto, "N")
    Conn.Execute Sql
    
'    If DevuelveValor("select ctrstock from sartic where codartic = " & DBSet(IdProducto, "N")) = 1 Then
'        SQL = "update sartic set " & _
'              "  canstock = canstock - " & DBSet(c_Cantidad, "N") & _
'              " where codartic = " & DBSet(IdProducto, "N")
'        Conn.Execute SQL
'    End If
    
    
    ' insertamos en la tabla de albaranes
    Codclave = SugerirCodigoSiguienteStr("scaalb", "codclave")
    
    forpa = ""
    forpa = DevuelveDesdeBDNew(cPTours, "sforpa", "codforpa", "forpaalvic", IdTipoPago, "N")
    
    '[Monica]30/11/2011 añadida segunda condicion
    If Trim(NumFactura) <> "" And InStr(1, Tarjeta, "Z") <> 0 Then
        codsoc = DevuelveDesdeBDNew(cPTours, "ssocio", "codsocio", "nifsocio", NifCliente, "T")
        If Mid(Tarjeta, 1, 4) = "****" Or Trim(Tarjeta) = "" Then
            Tarjeta = codsoc
            
        Else '[Monica]07/02/2011 buscamos la tarjeta que corresponda para meter pq me viene Z
            If codsoc >= 900000 Then
                Tarjeta = DevuelveValor("select numtarje from starje where codsocio= " & DBSet(codsoc, "N"))
            Else
                ' el socio es inferior a 900000 miro si hay tarjeta dependiendo del producto
                Dim TipArtic As Integer
                TipArtic = DevuelveValor("select tipogaso from sartic where codartic = " & DBSet(IdProducto, "N"))
                If TipArtic = 3 Then ' si el articulo es gasoleo bonificado
                    Tarjeta = DevuelveValor("select numtarje from starje where codsocio= " & DBSet(codsoc, "N") & " and tiptarje = 1")
                    
                    '[Monica]22/11/2011: si no tiene tarjeta de gasoleo bonificado le meto la primera tarjeta que encuentre
                    If Tarjeta = "0" Then
                        Tarjeta = DevuelveValor("select numtarje from starje where codsocio = " & DBSet(codsoc, "N"))
                    End If
                    
                Else
                    Tarjeta = DevuelveValor("select numtarje from starje where codsocio= " & DBSet(codsoc, "N") & " and tiptarje = 0")
                End If
            End If
            '[Monica]07/02/2011 hasta aqui
        End If
        'fechahora--> txtcodigo(0).Text & " " & Time
        
        Sql = "INSERT INTO scaalb (codclave, codsocio, numtarje, numalbar, fecalbar, horalbar, " & _
              "codturno, codartic, cantidad, preciove, importel, codforpa, matricul, codtraba, " & _
              "numfactu, numlinea, kilometros) VALUES (" & DBSet(Codclave, "T") & "," & DBSet(codsoc, "N") & "," & DBSet(Tarjeta, "N") & "," & _
               DBSet(NumAlbaran, "T") & "," & DBSet(txtCodigo(0).Text, "F") & "," & DBSet(txtCodigo(0).Text & " " & Hora, "FH") & "," & DBSet(Turno, "N") & "," & _
               DBSet(IdProducto, "N") & "," & DBSet(c_Cantidad, "N") & "," & DBSet(c_Precio, "N") & "," & _
               DBSet(c_Importe, "N") & "," & DBSet(forpa, "N") & "," & DBSet(Matricula, "T") & "," & DBSet(IdVendedor, "N") & ","
    
        numlinea = SugerirCodigoSiguienteStr("scaalb", "numlinea", "numfactu = " & DBSet(NumFactura, "N"))
        Sql = Sql & DBSet(NumFactura, "N") & "," & DBSet(numlinea, "N") & ","
        
        '[monica]24/06/2013: añadimos los kilometros
        Sql = Sql & DBSet(Round2(ComprobarCero(Trim(Kilometros)) / 100, 0), "N", "S") & ")"
        
        
        Conn.Execute Sql
        
        If c_Descuento <> 0 Then
            Sql = "update sartic set preventa = " & DBSet(c_Precio2, "N") & _
                  " where codartic = " & DBSet(IdProductoDes, "N")
            Conn.Execute Sql
            
            Codclave = SugerirCodigoSiguienteStr("scaalb", "codclave")
           
            Sql = "INSERT INTO scaalb (codclave, codsocio, numtarje, numalbar, fecalbar, horalbar, " & _
                  "codturno, codartic, cantidad, preciove, importel, codforpa, matricul, codtraba, " & _
                  "numfactu, numlinea) VALUES (" & DBSet(Codclave, "T") & "," & DBSet(codsoc, "N") & "," & DBSet(codsoc, "N") & "," & _
                   DBSet(NumAlbaran, "T") & "," & DBSet(txtCodigo(0).Text, "F") & "," & DBSet(txtCodigo(0).Text & " " & Hora, "FH") & "," & DBSet(Turno, "N") & "," & _
                   DBSet(IdProductoDes, "N") & "," & DBSet(c_Cantidad, "N") & "," & DBSet(c_Precio2, "N") & "," & _
                   DBSet(c_Importe2, "N") & "," & DBSet(forpa, "N") & "," & DBSet(Matricula, "T") & "," & DBSet(IdVendedor, "N") & ","
        
            numlinea = numlinea + 1
            Sql = Sql & DBSet(NumFactura, "N") & "," & DBSet(numlinea, "N") & ")"
        
            Conn.Execute Sql
        End If
        
    Else
        '[Monica]30/11/2010
        If Trim(NumFactura) <> "" Then
            codsoc = DevuelveDesdeBDNew(cPTours, "starje", "codsocio", "numtarje", Tarjeta, "N")
            If Mid(Tarjeta, 1, 4) = "****" Or Trim(Tarjeta) = "" Then
                Tarjeta = codsoc
            End If
            'fechahora--> txtcodigo(0).Text & " " & Time
            
            Sql = "INSERT INTO scaalb (codclave, codsocio, numtarje, numalbar, fecalbar, horalbar, " & _
                  "codturno, codartic, cantidad, preciove, importel, codforpa, matricul, codtraba, " & _
                  "numfactu, numlinea, kilometros) VALUES (" & DBSet(Codclave, "T") & "," & DBSet(codsoc, "N") & "," & DBSet(Tarjeta, "N") & "," & _
                   DBSet(NumAlbaran, "T") & "," & DBSet(txtCodigo(0).Text, "F") & "," & DBSet(txtCodigo(0).Text & " " & Hora, "FH") & "," & DBSet(Turno, "N") & "," & _
                   DBSet(IdProducto, "N") & "," & DBSet(c_Cantidad, "N") & "," & DBSet(c_Precio, "N") & "," & _
                   DBSet(c_Importe, "N") & "," & DBSet(forpa, "N") & "," & DBSet(Matricula, "T") & "," & DBSet(IdVendedor, "N") & ","
        
            numlinea = SugerirCodigoSiguienteStr("scaalb", "numlinea", "numfactu = " & DBSet(NumFactura, "N"))
            Sql = Sql & DBSet(NumFactura, "N") & "," & DBSet(numlinea, "N") & ","
            
            '[monica]24/06/2013: añadimos los kilometros
            Sql = Sql & DBSet(Round2(ComprobarCero(Trim(Kilometros)) / 100, 0), "N", "S") & ")"
            
            Conn.Execute Sql
            
            If c_Descuento <> 0 Then
                Sql = "update sartic set preventa = " & DBSet(c_Precio2, "N") & _
                      " where codartic = " & DBSet(IdProductoDes, "N")
                Conn.Execute Sql
                
                Codclave = SugerirCodigoSiguienteStr("scaalb", "codclave")
               
                Sql = "INSERT INTO scaalb (codclave, codsocio, numtarje, numalbar, fecalbar, horalbar, " & _
                      "codturno, codartic, cantidad, preciove, importel, codforpa, matricul, codtraba, " & _
                      "numfactu, numlinea) VALUES (" & DBSet(Codclave, "T") & "," & DBSet(codsoc, "N") & "," & DBSet(Tarjeta, "N") & "," & _
                       DBSet(NumAlbaran, "T") & "," & DBSet(txtCodigo(0).Text, "F") & "," & DBSet(txtCodigo(0).Text & " " & Hora, "FH") & "," & DBSet(Turno, "N") & "," & _
                       DBSet(IdProductoDes, "N") & "," & DBSet(c_Cantidad, "N") & "," & DBSet(c_Precio2, "N") & "," & _
                       DBSet(c_Importe2, "N") & "," & DBSet(forpa, "N") & "," & DBSet(Matricula, "T") & "," & DBSet(IdVendedor, "N") & ","
            
                numlinea = numlinea + 1
                Sql = Sql & DBSet(NumFactura, "N") & "," & DBSet(numlinea, "N") & ")"
            
                Conn.Execute Sql
            End If
        
        Else
            CodigoCliente = DevuelveDesdeBDNew(cPTours, "starje", "codsocio", "numtarje", Tarjeta, "N")
            Sql = "INSERT INTO scaalb (codclave, codsocio, numtarje, numalbar, fecalbar, horalbar, " & _
                  "codturno, codartic, cantidad, preciove, importel, codforpa, matricul, codtraba, " & _
                  "numfactu, numlinea, kilometros) VALUES (" & DBSet(Codclave, "T") & "," & DBSet(CodigoCliente, "N") & "," & DBSet(Tarjeta, "N") & "," & _
                   DBSet(NumAlbaran, "T") & "," & DBSet(txtCodigo(0).Text, "F") & "," & DBSet(txtCodigo(0).Text & " " & Hora, "FH") & "," & DBSet(Turno, "N") & "," & _
                   DBSet(IdProducto, "N") & "," & DBSet(c_Cantidad, "N") & "," & DBSet(c_Precio, "N") & "," & _
                   DBSet(c_Importe, "N") & "," & DBSet(forpa, "N") & "," & DBSet(Matricula, "T") & "," & DBSet(IdVendedor, "N") & ","
            Sql = Sql & "0,0,"
            
            '[monica]24/06/2013: añadimos los kilometros
            Sql = Sql & DBSet(Round2(ComprobarCero(Trim(Kilometros)) / 100, 0), "N", "S") & ")"
                    
            
            Conn.Execute Sql
            
            If c_Descuento <> 0 Then
                Sql = "update sartic set preventa = " & DBSet(c_Precio2, "N") & _
                      " where codartic = " & DBSet(IdProductoDes, "N")
                Conn.Execute Sql
                
                Codclave = SugerirCodigoSiguienteStr("scaalb", "codclave")
                
                Sql = "INSERT INTO scaalb (codclave, codsocio, numtarje, numalbar, fecalbar, horalbar, " & _
                      "codturno, codartic, cantidad, preciove, importel, codforpa, matricul, codtraba, " & _
                      "numfactu, numlinea) VALUES (" & DBSet(Codclave, "T") & "," & DBSet(CodigoCliente, "N") & "," & DBSet(Tarjeta, "N") & "," & _
                       DBSet(NumAlbaran, "T") & "," & DBSet(txtCodigo(0).Text, "F") & "," & DBSet(txtCodigo(0).Text & " " & Hora, "FH") & "," & DBSet(Turno, "N") & "," & _
                       DBSet(IdProductoDes, "N") & "," & DBSet(c_Cantidad, "N") & "," & DBSet(c_Precio2, "N") & "," & _
                       DBSet(c_Importe2, "N") & "," & DBSet(forpa, "N") & "," & DBSet(Matricula, "T") & "," & DBSet(IdVendedor, "N") & ","
                Sql = Sql & "0,0)"
            
                Conn.Execute Sql
            End If
        End If
    End If
 
    
    
eInsertarLineaRib:
    If Err.Number <> 0 Then
        InsertarLineaRib = False
        MsgBox "Error en Insertar Linea " & Err.Description, vbExclamation
    End If
End Function
            
            
            
            
            
            
            
            
Private Function InsertarRecaudacion(cad As String) As Boolean
Dim forpa As String
Dim Importe As String
Dim Sql As String
Dim vImporte As String
Dim IdTipoPago As String
Dim Existe As String

    On Error GoTo eInsertarRecaudacion

    InsertarRecaudacion = True
'    forpa = Mid(cad, 2, 2)
'    Importe = Mid(cad, 14, 8) & "," & Mid(cad, 22, 2)
    Importe = Mid(cad, 668, 18)
    
    If Trim(Importe) = "" Then
        Exit Function
    Else
        If CCur(Importe) = 0 Then Exit Function
    End If
    
    IdTipoPago = Mid(cad, 784, 10)
    vImporte = Round2(CCur(Importe) / 100, 2)

    forpa = ""
    forpa = DevuelveDesdeBDNew(cPTours, "sforpa", "codforpa", "forpaalvic", IdTipoPago, "N")
    
    If CCur(vImporte) <> 0 Then
        Existe = ""
        Existe = DevuelveDesdeBDNew(cPTours, "srecau", "codforpa", "fechatur", txtCodigo(0).Text, "F", , "codturno", txtCodigo(1).Text, "N", "codforpa", forpa, "N")
        If Existe = "" Then
            Sql = "insert into srecau (fechatur, codturno, codforpa, importel, intconta) values (" & _
                  DBSet(txtCodigo(0).Text, "F") & "," & DBSet(txtCodigo(1).Text, "N") & "," & _
                  DBSet(CInt(forpa), "N") & "," & DBSet(vImporte, "N") & ",0)"
        Else
            Sql = "update srecau set importel = importel + " & DBSet(vImporte, "N")
            Sql = Sql & " where fechatur = " & DBSet(txtCodigo(0).Text, "F")
            Sql = Sql & " and codturno = " & DBSet(txtCodigo(1).Text, "N")
            Sql = Sql & " and codforpa = " & DBSet(forpa, "N")
        End If
        Conn.Execute Sql
    End If
eInsertarRecaudacion:
    If Err.Number <> 0 Then
        InsertarRecaudacion = False
        MsgBox "Error en Insertar Recaudacion en " & Err.Description, vbExclamation
    End If
    
End Function

Private Function InsertarSalida(cad As String) As Boolean
Dim TipMov As String
Dim Importe As Currency
Dim Sql As String
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
        Sql = "insert into srecau (fechatur, codturno, codforpa, importel, intconta) values (" & _
              DBSet(txtCodigo(0).Text, "F") & "," & DBSet(txtCodigo(1).Text, "N") & "," & _
              "99, " & DBSet(Importe, "N") & ",0)"
              
        Conn.Execute Sql
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
Dim Precio As String
Dim Importe As String
Dim Sql As String
Dim NumLin As Long
Dim cWhere As String


    codpro = Mid(cad, 35, 2)
    cantidad = Mid(cad, 54, 6) & "," & Mid(cad, 60, 2)
    Precio = Mid(cad, 42, 2) & "," & Mid(cad, 44, 2)
    Importe = Mid(cad, 47, 5) & "," & Mid(cad, 52, 2)
    
    Sql = ""
    Sql = DevuelveDesdeBDNew(cPTours, "sturno", "codturno", "fechatur", txtCodigo(0).Text, "F", , "codturno", txtCodigo(1).Text, "N", "codartic", codpro, "N")
    If Sql = "" Then
    
        cWhere = "fechatur=" & DBSet(txtCodigo(0).Text, "F") & " and codturno = " & DBSet(txtCodigo(1).Text, "N")
        NumLin = CLng(SugerirCodigoSiguienteStr("sturno", "numlinea", cWhere))
        'insertamos
        Sql = "INSERT INTO sturno (fechatur, codturno, numlinea, tiporegi, numtanqu, nummangu, " & _
              " codartic, litrosve, importel, containi, contafin, tipocred) VALUES (" & _
              DBSet(txtCodigo(0).Text, "F") & "," & DBSet(txtCodigo(1).Text, "N") & "," & DBSet(NumLin, "N") & ",2,1,1," & _
              DBSet(codpro, "N") & "," & DBSet(cantidad, "N") & "," & DBSet(Importe, "N") & ",0,0,0)"
              
        Conn.Execute Sql
    Else
        'actualizamos
        Sql = "UPDATE sturno SET importel = importel + " & DBSet(Importe, "N") & ", litrosve = litrosve +  " & DBSet(cantidad, "N") & " WHERE fechatur = " & _
              DBSet(txtCodigo(0).Text, "F") & " AND codturno = " & DBSet(txtCodigo(1).Text, "N") & " AND codartic = " & _
              DBSet(codpro, "N")
              
        Conn.Execute Sql
    End If
End Sub

Private Function FicheroCorrecto(Tipo As String) As Boolean
Dim fic As String
Dim fec As Date
    fec = CDate(txtCodigo(0).Text)
    
    FicheroCorrecto = (UCase(NombreFichero(Me.CommonDialog1.FileName)) = UCase(("VENTAS" & Format(Day(fec), "00") & Format(Month(fec), "00") & Format(Year(fec) - 2000, "00") & "-" & Format(txtCodigo(1).Text, "00") & ".txt")))
    
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
Dim Sql As String
    Sql = "delete from tmpinformes where codusu = " & vSesion.Codigo
    
    Conn.Execute Sql
End Sub

'Public Function CrearTMPCargas(cadTABLA As String, cadwhere As String) As Boolean
''Crea una temporal donde inserta las cargas del fichero para trabajar con ellas
'Dim sql As String
'
'    On Error GoTo ECrear
'
'    CrearTMPCargas = False
'
'    sql = "CREATE TEMPORARY TABLE tmpcargas ( "
'    sql = sql & "base integer,"
'    sql = sql & "nombrebase varchar(50),"
'    sql = sql & "turno integer,"
'    sql = sql & "numalbaran varchar(20),"
'    sql = sql & "numfactura varchar(20),"
'    sql = sql & "idvendedor integer,"
'    sql = sql & "nombrevendedor varchar(50),"
'    sql = sql & "fechahora datetime,"
'    sql = sql & "codigocliente varchar(20),"
'    sql = sql & "nombrecliente varchar(75),"
'    sql = sql & "matricula varchar(20),"
'    sql = sql & "codigoproducto varchar(20),"
'    sql = sql & "surtidor integer,"
'    sql = sql & "manguera integer,"
'    sql = sql & "preciolitro decimal(18,5),"
'    sql = sql & "cantidad decimal(18,2),"
'    sql = sql & "importe decimal(18,2),"
'    sql = sql & "idtipopago integer,"
'    sql = sql & "descrtipopago varchar(25),"
'    sql = sql & "codigotipopago varchar(15),"
'    sql = sql & "nifcliente varchar(20))"
'    Conn.Execute sql
'
'    CrearTMPCargas = True
'
'ECrear:
'     If Err.Number <> 0 Then
'        CrearTMPCargas = False
'        'Borrar la tabla temporal
'        sql = " DROP TABLE IF EXISTS tmpcargas;"
'        Conn.Execute sql
'    End If
'End Function


'Public Sub BorrarTMPCargas()
'On Error Resume Next
'
'    Conn.Execute " DROP TABLE IF EXISTS tmpcargas;"
'    If Err.Number <> 0 Then Err.Clear
'End Sub

Private Function InsertarRecaudacionAlz(fic As String) As Boolean
Dim NF As Long
Dim i As Long
Dim longitud As Long

Dim forpa As String
Dim Importe As String
Dim TipoMov As String
Dim Sql As String
Dim IdTipoPago As String
Dim Existe As String
Dim Forpa1 As String
Dim Importe1 As Currency

Dim Fic1 As String

    On Error GoTo eInsertarRecaudacionAlz

    InsertarRecaudacionAlz = True
    
    '****** PROCESAMOS EL FICHERO DE TOTALES
    If Dir(fic) = "" Then
        If MsgBox("No existe el fichero de totales. ¿Desea continuar?", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
            InsertarRecaudacionAlz = False
            Exit Function
        End If
    Else
        NF = FreeFile
    
        Open fic For Input As #NF '
        
        Line Input #NF, cad
        i = 0
        
        lblProgres(0).Caption = "Procesando Fichero: " & fic
        longitud = FileLen(fic)
        
        Pb1.visible = True
        Me.Pb1.Max = longitud
        Me.Refresh
        Me.Pb1.Value = 0
        
        While Not EOF(NF)
            i = i + 1
            
            Me.Pb1.Value = Me.Pb1.Value + Len(cad)
            lblProgres(1).Caption = "Linea " & i
            Me.Refresh
        
            forpa = Mid(cad, 71, 10)
            TipoMov = Mid(cad, 106, 10)
            Importe = Mid(cad, 141, 18)
            
            If CCur(forpa) <> 0 And CCur(TipoMov) = 0 And CCur(Importe) <> 0 Then
                Forpa1 = DevuelveDesdeBDNew(cPTours, "sforpa", "codforpa", "forpaalvic", forpa, "N")
                Importe1 = Round2(CCur(Importe) / 100000, 5)
            
                Sql = "insert into srecau (fechatur, codturno, codforpa, importel, intconta) values ("
                Sql = Sql & DBSet(txtCodigo(0).Text, "F")
                Sql = Sql & "," & DBSet(txtCodigo(1).Text, "N")
                Sql = Sql & "," & DBSet(CInt(Forpa1), "N")
                Sql = Sql & "," & DBSet(Importe1, "N") & ",0)"
                
                Conn.Execute Sql
            End If
            
            Line Input #NF, cad
        Wend
        If cad <> "" Then
            Me.Pb1.Value = Me.Pb1.Value + Len(cad)
            lblProgres(1).Caption = "Linea " & i
            Me.Refresh
        
            forpa = Mid(cad, 71, 10)
            TipoMov = Mid(cad, 106, 10)
            Importe = Mid(cad, 141, 18)
            
            If CCur(forpa) <> 0 And CCur(TipoMov) = 0 And CCur(Importe) <> 0 Then
                Forpa1 = DevuelveDesdeBDNew(cPTours, "sforpa", "codforpa", "forpaalvic", forpa, "N")
                Importe1 = Round2(CCur(Importe) / 100000, 5)
            
                Sql = "insert into srecau (fechatur, codturno, codforpa, importel, intconta) values ("
                Sql = Sql & DBSet(txtCodigo(0).Text, "F")
                Sql = Sql & "," & DBSet(txtCodigo(1).Text, "N")
                Sql = Sql & "," & DBSet(CInt(Forpa1), "N")
                Sql = Sql & "," & DBSet(Importe1, "N") & ",0)"
                
                Conn.Execute Sql
            End If
        End If
        Close #NF
    End If
        
    '****** PROCESAMOS EL FICHERO DE CAJA
  
    NF = FreeFile

    Fic1 = Replace(fic, "totales", "caja")
    
    If Dir(Fic1) = "" Then
        If MsgBox("No existe el fichero de cajas. ¿Desea continuar?", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
            InsertarRecaudacionAlz = False
            Exit Function
        End If
    Else
        Open Fic1 For Input As #NF '
    
        Line Input #NF, cad
    
        i = 0
    
        lblProgres(0).Caption = "Procesando Fichero: " & Fic1
        longitud = FileLen(Fic1)
    
        Pb1.visible = True
        Me.Pb1.Max = longitud
        Me.Refresh
        Me.Pb1.Value = 0
        While Not EOF(NF)
            i = i + 1
    
            Me.Pb1.Value = Me.Pb1.Value + Len(cad)
            lblProgres(1).Caption = "Linea " & i
            Me.Refresh
    
            TipoMov = Mid(cad, 254, 10)
            Importe = Mid(cad, 236, 18)
    
            If CCur(TipoMov) = 1 And CCur(Importe) <> 0 Then
                Importe1 = Round2(CCur(Importe) / 100000, 5)
    
                Sql = "select count(*) from srecau where fechatur = " & DBSet(txtCodigo(0).Text, "F")
                Sql = Sql & " and codturno = " & DBSet(txtCodigo(1).Text, "N")
                Sql = Sql & " and codforpa = 99"
                If TotalRegistros(Sql) = 0 Then
                    Sql = "insert into srecau (fechatur, codturno, codforpa, importel, intconta) values ("
                    Sql = Sql & DBSet(txtCodigo(0).Text, "F")
                    Sql = Sql & "," & DBSet(txtCodigo(1).Text, "N")
                    Sql = Sql & ",99" ' Introducimos a piñon la forpa 99
                    Sql = Sql & "," & DBSet(Importe1, "N") & ",0)"
                Else
                    Sql = "update srecau set importel = importel + " & DBSet(Importe1, "N")
                    Sql = Sql & " where fechatur = " & DBSet(txtCodigo(0).Text, "F")
                    Sql = Sql & " and codturno = " & DBSet(txtCodigo(1).Text, "N")
                    Sql = Sql & " and codforpa = 99"
                End If
    
                Conn.Execute Sql
            End If
    
            '++monica: 09/05/08 introducimos tambien seguridad
            If CCur(TipoMov) = 3 And CCur(Importe) <> 0 Then
                Importe1 = Round2(CCur(Importe) / 100000, 5)
    
                Sql = "select count(*) from srecau where fechatur = " & DBSet(txtCodigo(0).Text, "F")
                Sql = Sql & " and codturno = " & DBSet(txtCodigo(1).Text, "N")
                Sql = Sql & " and codforpa = 97"
                If TotalRegistros(Sql) = 0 Then
                    Sql = "insert into srecau (fechatur, codturno, codforpa, importel, intconta) values ("
                    Sql = Sql & DBSet(txtCodigo(0).Text, "F")
                    Sql = Sql & "," & DBSet(txtCodigo(1).Text, "N")
                    Sql = Sql & ",97" ' Introducimos a piñon la forpa 97
                    Sql = Sql & "," & DBSet(Importe1, "N") & ",0)"
                Else
                    Sql = "update srecau set importel = importel + " & DBSet(Importe1, "N")
                    Sql = Sql & " where fechatur = " & DBSet(txtCodigo(0).Text, "F")
                    Sql = Sql & " and codturno = " & DBSet(txtCodigo(1).Text, "N")
                    Sql = Sql & " and codforpa = 97"
                End If
    
                Conn.Execute Sql
            End If
    
    
            Line Input #NF, cad
        Wend
        If cad <> "" Then
            i = i + 1
    
            Me.Pb1.Value = Me.Pb1.Value + Len(cad)
            lblProgres(1).Caption = "Linea " & i
            Me.Refresh
    
            TipoMov = Mid(cad, 254, 10)
            Importe = Mid(cad, 236, 18)
    
            If CCur(TipoMov) = 1 And CCur(Importe) <> 0 Then
                Importe1 = Round2(CCur(Importe) / 100000, 5)
    
                Sql = "select count(*) from srecau where fechatur = " & DBSet(txtCodigo(0).Text, "F")
                Sql = Sql & " and codturno = " & DBSet(txtCodigo(1).Text, "N")
                Sql = Sql & " and codforpa = 99"
                If TotalRegistros(Sql) = 0 Then
                    Sql = "insert into srecau (fechatur, codturno, codforpa, importel, intconta) values ("
                    Sql = Sql & DBSet(txtCodigo(0).Text, "F")
                    Sql = Sql & "," & DBSet(txtCodigo(1).Text, "N")
                    Sql = Sql & ",99" ' Introducimos a piñon la forpa 99
                    Sql = Sql & "," & DBSet(Importe1, "N") & ",0)"
                Else
                    Sql = "update srecau set importel = importel + " & DBSet(Importe1, "N")
                    Sql = Sql & " where fechatur = " & DBSet(txtCodigo(0).Text, "F")
                    Sql = Sql & " and codturno = " & DBSet(txtCodigo(1).Text, "N")
                    Sql = Sql & " and codforpa = 99"
                End If
    
                Conn.Execute Sql
            End If
            
            '++monica: 09/05/08 incluimos seguridad
            If CCur(TipoMov) = 3 And CCur(Importe) <> 0 Then
                Importe1 = Round2(CCur(Importe) / 100000, 5)
    
                Sql = "select count(*) from srecau where fechatur = " & DBSet(txtCodigo(0).Text, "F")
                Sql = Sql & " and codturno = " & DBSet(txtCodigo(1).Text, "N")
                Sql = Sql & " and codforpa = 97"
                If TotalRegistros(Sql) = 0 Then
                    Sql = "insert into srecau (fechatur, codturno, codforpa, importel, intconta) values ("
                    Sql = Sql & DBSet(txtCodigo(0).Text, "F")
                    Sql = Sql & "," & DBSet(txtCodigo(1).Text, "N")
                    Sql = Sql & ",97" ' Introducimos a piñon la forpa 97
                    Sql = Sql & "," & DBSet(Importe1, "N") & ",0)"
                Else
                    Sql = "update srecau set importel = importel + " & DBSet(Importe1, "N")
                    Sql = Sql & " where fechatur = " & DBSet(txtCodigo(0).Text, "F")
                    Sql = Sql & " and codturno = " & DBSet(txtCodigo(1).Text, "N")
                    Sql = Sql & " and codforpa = 97"
                End If
    
                Conn.Execute Sql
            End If
        
        Close #NF
        
        End If
    End If
eInsertarRecaudacionAlz:
    If Err.Number <> 0 Then
        InsertarRecaudacionAlz = False
        MsgBox "Error en Insertar Recaudacion en " & Err.Description, vbExclamation
    End If
End Function


Private Function InsertarLineaTurnoNew(fic As String) As Boolean
Dim NF As Long
Dim i As Long
Dim longitud As Long


Dim codpro As String
Dim cantidad As String
Dim Precio As String
Dim Importe As String
Dim Sql As String
Dim NumLin As Long
Dim cWhere As String

Dim Surtidor As String
Dim Manguera As String
Dim Inicial As String
Dim Final As String
Dim vInicial As Currency
Dim vFinal As Currency

    On Error GoTo eInsertarLineaTurnoNew

    InsertarLineaTurnoNew = True

    If Dir(fic) = "" Then
        If MsgBox("No existe el fichero de totales. ¿Desea continuar?", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
            InsertarLineaTurnoNew = False
            Exit Function
        End If
    Else

        NF = FreeFile
        
        
        '****** PROCESAMOS EL FICHERO DE TOTALES
        
        Open fic For Input As #NF '
        
        Line Input #NF, cad
        i = 0
        
        lblProgres(0).Caption = "Procesando Fichero: " & fic
        longitud = FileLen(fic)
        
        Pb1.visible = True
        Me.Pb1.Max = longitud
        Me.Refresh
        Me.Pb1.Value = 0
        
        While Not EOF(NF)
            i = i + 1
            
            Me.Pb1.Value = Me.Pb1.Value + Len(cad)
            lblProgres(1).Caption = "Linea " & i
            Me.Refresh
        
            codpro = Mid(cad, 162, 10)
            Surtidor = Mid(cad, 71, 10)
            Manguera = Mid(cad, 91, 10)
            Inicial = Mid(cad, 115, 18)
            Final = Mid(cad, 133, 18)
            vInicial = Round2(CCur(Inicial) / 100, 2)
            vFinal = Round2(CCur(Final / 100), 2)
            
            If CCur(vInicial) <> 0 And CCur(vFinal) <> 0 Then
                
                cWhere = "fechatur=" & DBSet(txtCodigo(0).Text, "F") & " and codturno = " & DBSet(txtCodigo(1).Text, "N")
                NumLin = CLng(SugerirCodigoSiguienteStr("sturno", "numlinea", cWhere))
                'insertamos
                Sql = "INSERT INTO sturno (fechatur, codturno, numlinea, tiporegi, numtanqu, nummangu, " & _
                      " codartic, litrosve, importel, containi, contafin, tipocred) VALUES (" & _
                      DBSet(txtCodigo(0).Text, "F") & "," & DBSet(txtCodigo(1).Text, "N") & "," & DBSet(NumLin, "N") & ",0," & _
                      DBSet(Surtidor, "N") & "," & DBSet(Manguera, "N") & "," & _
                      DBSet(codpro, "N") & ",0,0," & DBSet(vInicial, "N") & "," & DBSet(vFinal, "N") & ",0)"
                  
                
                Conn.Execute Sql
            End If
            
            Line Input #NF, cad
        Wend
        Close #NF
    End If
eInsertarLineaTurnoNew:
    If Err.Number <> 0 Then
        InsertarLineaTurnoNew = False
        MsgBox "Error en Insertar Turno en " & Err.Description, vbExclamation
    End If
End Function

'Private Function ProcesarFicheroCaja(fic As String) As Boolean
'Dim nf As Long
'Dim i As Long
'Dim longitud As Long
'
'Dim Forpa As String
'Dim Importe As String
'Dim TipoMov As String
'Dim Sql As String
'Dim IdTipoPago As String
'Dim Existe As String
'Dim Forpa1 As String
'Dim Importe1 As Currency
'
'    On Error GoTo eProcesarFicheroCaja
'
'    ProcesarFicheroCaja = True
'    nf = FreeFile
'
'    Fic1 = Replace(fic, "totales", "caja")
'
'    Open Fic1 For Input As #nf '
'
'    Line Input #nf, cad
'
'    i = 0
'
'    lblProgres(0).Caption = "Procesando Fichero: " & Fic1
'    longitud = FileLen(Fic1)
'
'    Pb1.visible = True
'    Me.Pb1.Max = longitud
'    Me.Refresh
'    Me.Pb1.Value = 0
'    While Not EOF(nf)
'        i = i + 1
'
'        Me.Pb1.Value = Me.Pb1.Value + Len(cad)
'        lblProgres(1).Caption = "Linea " & i
'        Me.Refresh
'
'        TipoMov = Mid(cad, 254, 10)
'        Importe = Mid(cad, 236, 18)
'
'        If CCur(TipoMov) = 1 And CCur(Importe) <> 0 Then
'            Importe1 = Round2(CCur(Importe) / 100000, 5)
'
'            Sql = "select count(*) from srecau where fechatur = " & DBSet(txtcodigo(0).Text, "F")
'            Sql = Sql & " and codturno = " & DBSet(txtcodigo(1).Text, "N")
'            Sql = Sql & " and codforpa = 99"
'            If TotalRegistros(Sql) = 0 Then
'                Sql = "insert into srecau (fechatur, codturno, codforpa, importel, intconta) values ("
'                Sql = Sql & DBSet(txtcodigo(0).Text, "F")
'                Sql = Sql & "," & DBSet(txtcodigo(1).Text, "N")
'                Sql = Sql & ",99" ' Introducimos a piñon la forpa 99
'                Sql = Sql & "," & DBSet(Importe1, "N") & ",0)"
'            Else
'                Sql = "update srecau set importe = importe + " & DBSet(Importe1, "N")
'                Sql = Sql & " where fechatur = " & DBSet(txtcodigo(0).Text, "F")
'                Sql = Sql & " and codturno = " & DBSet(txtcodigo(1).Text, "N")
'                Sql = Sql & " and codforpa = 99"
'            End If
'
'            Conn.Execute Sql
'        End If
'
'        Line Input #nf, cad
'    Wend
'    Close #nf
'
'eProcesarFicheroCaja:
'    If Err.Number <> 0 Then
'        ProcesarFicheroCaja = False
'    End If
'End Function


Private Function ProcesarFicheroCompras(nomFich As String) As Boolean
Dim NF As Long
Dim cad As String
Dim i As Integer
Dim longitud As Long
Dim RS As ADODB.Recordset
Dim Rs1 As ADODB.Recordset
Dim Numreg As Long
Dim Sql As String
Dim SQL1 As String
Dim total As Long
Dim v_cant As Currency
Dim v_impo As Currency
Dim v_prec As Currency
Dim b As Boolean
Dim NomFic As String
Dim MensError As String
    
    On Error GoTo eProcesarFicheroCompras


    NF = FreeFile
    
    Open nomFich For Input As #NF ' & "\BV" & Format(CDate(txtcodigo(0).Text), "ddmmyy") & "." & Format(txtcodigo(1).Text, "000") For Input As #NF
    
    Line Input #NF, cad
    i = 0
    
    lblProgres(0).Caption = "Procesando Fichero Compras: " & nomFich
    longitud = FileLen(nomFich)
    
    Pb1.visible = True
    Me.Pb1.Max = longitud
    Me.Refresh
    Me.Pb1.Value = 0
        
    ProcesarFicheroCompras = False
    
    BorrarTMP
    b = CrearTMP()
    If Not b Then
         Exit Function
    End If
        
    Conn.BeginTrans
        
    b = True
    MensError = "Error Insertando Linea de Albarán de Compras:"
    While Not EOF(NF) And b
        i = i + 1
        
        Me.Pb1.Value = Me.Pb1.Value + Len(cad)
        lblProgres(1).Caption = "Linea " & i
        Me.Refresh
        
        b = InsertarLineaCompras(cad, MensError)
        
        Line Input #NF, cad
    Wend
    Close #NF
    
    If cad <> "" And b Then
        b = InsertarLineaCompras(cad, MensError)
        
        If b Then
            b = PasarTemporales()
        End If
    End If
    
    
eProcesarFicheroCompras:
    If Err.Number <> 0 Or Not b Then
        ProcesarFicheroCompras = False
        Conn.RollbackTrans
        MsgBox "No se ha realizado la importación del fichero de compras" & vbCrLf & vbCrLf & MensError, vbExclamation
    
        Pb1.visible = False
        lblProgres(0).Caption = ""
        lblProgres(1).Caption = ""
    Else
        ProcesarFicheroCompras = True
        Conn.CommitTrans
        
        Pb1.visible = False
        lblProgres(0).Caption = ""
        lblProgres(1).Caption = ""
    End If
End Function


Private Function InsertarLineaCompras(cad As String, ByRef MensError As String) As Boolean
Dim Sql As String

Dim c_Cantidad As Currency
Dim c_Importe As Currency
Dim c_Precio As Double
Dim c_PorcIva As Currency

Dim Mens As String
Dim numlinea As Long

Dim Albaran As String
Dim fechahora As String
Dim Proveedor As String
Dim NomProve As String
Dim IdProducto As String
Dim NomArtic As String
Dim cantidad As String
Dim PorcIva As String
Dim Precio As String
Dim Importe As String

Dim forpa As String
Dim Banco As String

Dim DomProv  As String
Dim CPostalProv As String
Dim PobProv As String
Dim ProProv As String
Dim NIFProv As String
Dim TelProv As String
Dim vProve As CProveedor

Dim Fecha As String
Dim TipoIva As String

Dim Familia As Integer
Dim Codmacta As String
Dim Codmaccl As String
Dim Rsf As ADODB.Recordset

    On Error GoTo eInsertarLinea

    InsertarLineaCompras = False
    

    Albaran = Trim(Mid(cad, 92, 15))
    ' si la longitud es mayor de 10 cogemos los 10 ultimos caracteres
    If Len(Albaran) > 10 Then Albaran = Mid(Albaran, Len(Albaran) - 9, 10)
    fechahora = Mid(cad, 122, 14)
    Proveedor = Mid(cad, 136, 10)
    NomProve = Mid(cad, 146, 40)
    IdProducto = Mid(cad, 580, 15)
    NomArtic = Mid(cad, 333, 25)
    cantidad = Mid(cad, 453, 18)
    PorcIva = Mid(cad, 471, 18)
    Precio = Mid(cad, 543, 18)
    Importe = Mid(cad, 561, 18)
    
    Fecha = Mid(fechahora, 7, 2) & "/" & Mid(fechahora, 5, 2) & "/" & Mid(fechahora, 1, 4)
    fechahora = Mid(fechahora, 1, 4) & "-" & Mid(fechahora, 5, 2) & "-" & Mid(fechahora, 7, 2) & " " & Mid(fechahora, 9, 2) & ":" & Mid(fechahora, 11, 2) & ":" & Mid(fechahora, 13, 2)

    c_Cantidad = Round2(CCur(cantidad) / 100, 2)
    c_Importe = Round2(CCur(Importe) / 100, 2)
    c_Precio = Round2(CDbl(Precio) / 100000, 5)


'    'VRS:4.0.1(0) actualizamos el precio de articulo cuando pasamos las temporales
'    Sql = "update sartic set ultpreci = " & DBSet(c_Precio, "N") & _
'          ", ultfecha = " & DBSet(Fecha, "F") & _
'          ", canstock = canstock + " & DBSet(c_Cantidad, "N") & _
'          " where codartic = " & DBSet(IdProducto, "N")
'    Conn.Execute Sql
    
    
    ' Comprobamos que existe el proveedor y si no lo creamos con el domicilio automático
    Set vProve = New CProveedor
    If TotalRegistros("select count(*) from proveedor where codprove = " & DBSet(Proveedor, "N")) <> 0 Then
        vProve.LeerDatos (Proveedor)
        '#### Leer estos datos de la tabla scaalpr y no de sprove
        NomProve = vProve.Nombre
        DomProv = vProve.Domicilio
        CPostalProv = vProve.CPostal
        PobProv = vProve.POBLACION
        ProProv = vProve.Provincia
        NIFProv = vProve.NIF
        TelProv = vProve.TfnoAdmon
        
        forpa = vProve.ForPago
        Banco = vProve.BancoPropio
        
    Else
        forpa = DevuelveValor("select min(codforpa) from sforpa")
        Banco = DevuelveValor("select min(codbanpr) from sbanco")
        
        DomProv = "AUTOMATICO"
        CPostalProv = "46"
        PobProv = "A"
        ProProv = "A"
        NIFProv = "A"
        TelProv = ""
        
        Sql = "insert into proveedor (codprove,nomprove,nomcomer,domprove,codpobla,pobprove,proprove,nifprove,fecprove,codmacta,codforpa,codbanpr,fechamov) values ("
        Sql = Sql & DBSet(Proveedor, "N") & "," & DBSet(NomProve, "T") & "," & DBSet(NomProve, "T") & ","
        Sql = Sql & "'AUTOMATICO'," & DBSet(CPostalProv, "T") & "," & DBSet(PobProv, "T") & "," & DBSet(ProProv, "T") & ","
        Sql = Sql & DBSet(NIFProv, "T") & "," & DBSet(Fecha, "F") & "," & DBSet(vParamAplic.CtaFamDefecto, "T") & ","
        Sql = Sql & DBSet(forpa, "N") & "," & DBSet(Banco, "N") & "," & DBSet(Fecha, "F") & ")"
        
        Conn.Execute Sql
    End If
    
    ' Comprobamos que existe el articulo sino lo creamos con los datos basicos que tengamos
    If TotalRegistros("select count(*) from sartic where codartic = " & DBSet(IdProducto, "N")) = 0 Then
        c_PorcIva = Round2(CCur(PorcIva) / 100, 5)

        TipoIva = ""
        TipoIva = DevuelveDesdeBDNew(cConta, "tiposiva", "codigiva", "porceiva", DBSet(c_PorcIva, "N"), "N")
    
        '[Monica]15/12/2010: tenemos que comprobar los dos primeros digitos que son la familia
        ' si existe, si no crearla
        Familia = Mid(Format(IdProducto, "00000"), 1, 2)
        
        Sql = "select count(*) from sfamia where codfamia = " & DBSet(Familia, "N")
        If TotalRegistros(Sql) <> 0 Then
            Sql = "select * from sartic where codfamia = " & DBSet(Familia, "N")
            
            Set Rsf = New ADODB.Recordset
            Rsf.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
            Codmacta = ""
            Codmaccl = ""
            If Not Rsf.EOF Then
                Codmacta = DBLet(Rsf!Codmacta, "T")
                Codmaccl = DBLet(Rsf!Codmaccl, "T")
            End If
            Set Rsf = Nothing
        Else
            Sql = "insert into sfamia (codfamia,nomfamia,tipfamia) values (" & DBSet(Familia, "N") & ","
            Sql = Sql & "'AUTOMATICO',0)"
            
            Conn.Execute Sql
        End If
        
        Sql = "insert into sartic (codartic,nomartic,codfamia,codmacta,codmaccl,codigiva,canstock,preciopmp,ultpreci,ultfecha,ctrstock,ctacompr,artnuevo) values ("
        Sql = Sql & DBSet(IdProducto, "N") & "," & DBSet(NomArtic, "T")
        Sql = Sql & "," & DBSet(Familia, "N") & "," ' ",0," ' la famlia la marcada por el articulo
        Sql = Sql & DBSet(Codmacta, "T") & "," 'DBSet(vParamAplic.CtaFamDefecto, "T") & ","
        Sql = Sql & DBSet(Codmaccl, "T") & "," 'DBSet(vParamAplic.CtaFamDefecto, "T") & ","
        Sql = Sql & DBSet(TipoIva, "N") & "," & DBSet(c_Cantidad, "N") & ","
        Sql = Sql & DBSet(c_Precio, "N") & "," & DBSet(c_Precio, "N") & "," & DBSet(Fecha, "F") & "," & DBSet(vParamAplic.ControlStock, "N") & "," & DBSet(vParamAplic.CtaFamDefecto, "T")
        Sql = Sql & ",1)" ' lo marcamos como articulo nuevo
        
        Conn.Execute Sql
    Else
        ' si existe el articulo el nombre que vale es el que tengo grabado en arigasol
        NomArtic = DevuelveValor("select nomartic from sartic where codartic = " & DBSet(IdProducto, "N"))
    End If
    
    
    Sql = "select count(*) from tmpscaalp where numalbar = " & DBSet(Trim(Albaran), "T") & " and fechaalb = " & DBSet(Fecha, "F") & " and codprove = " & DBSet(Proveedor, "N")
    If TotalRegistros(Sql) = 0 Then
        Sql = "insert into tmpscaalp (numalbar,fechaalb,codprove,nomprove,domprove,codpobla,pobprove,proprove,"
        Sql = Sql & "nifprove,codforpa,dtoppago,dtognral,fecturno,codturno) values (" & DBSet(Albaran, "T") & "," & DBSet(Fecha, "F") & ","
        Sql = Sql & DBSet(Proveedor, "N") & "," & DBSet(NomProve, "T") & "," & DBSet(DomProv, "T") & "," & DBSet(CPostalProv, "T") & ","
        Sql = Sql & DBSet(PobProv, "T") & "," & DBSet(ProProv, "T") & "," & DBSet(NIFProv, "T") & "," & DBSet(forpa, "N") & ","
        Sql = Sql & "0,0," & DBSet(txtCodigo(0).Text, "F") & ","
        
        If txtCodigo(1).Text <> "" Then
            Sql = Sql & DBSet(txtCodigo(1).Text, "N") & ")"
        Else
            '[Monica]15/01/2013: en el caso de Ribarroja no meten el turno de traspaso, meto un 1 por defecto
            Sql = Sql & "1)"
        End If
            
    
        Conn.Execute Sql
    End If
    
    numlinea = DevuelveValor("select max(numlinea) + 1 from tmpslialp where numalbar = " & DBSet(Albaran, "T") & " and fechaalb = " & DBSet(Fecha, "F") & " and codprove = " & DBSet(Proveedor, "N"))
    If numlinea = 0 Then numlinea = 1
    Sql = "insert into tmpslialp (numalbar,fechaalb,codprove,numlinea,codartic,codalmac,nomartic,ampliaci,cantidad,precioar,dtoline1,dtoline2,importel,fechahora) values ("
    Sql = Sql & DBSet(Albaran, "T") & "," & DBSet(Fecha, "F") & "," & DBSet(Proveedor, "N") & "," & DBSet(numlinea, "N") & ","
    Sql = Sql & DBSet(IdProducto, "N") & ",1," ' el almacen siempre va a ser 1
    Sql = Sql & DBSet(NomArtic, "T") & "," & ValorNulo & ","
    Sql = Sql & DBSet(c_Cantidad, "N") & "," & DBSet(c_Precio, "N") & ",0,0," & DBSet(c_Importe, "N") & "," & DBSet(fechahora, "FH") & ")"
    
    Conn.Execute Sql
        
    InsertarLineaCompras = True
    Exit Function
    
eInsertarLinea:
    InsertarLineaCompras = False
    MensError = MensError & Err.Description
End Function


Private Function CrearTMP() As Boolean
' temporales de lineas para insertar posteriormente en scaalp y slialp
Dim Sql As String
    
    On Error GoTo ECrear
    
    CrearTMP = False
    
    'tabla temporal con la que cargaremos: scaalp
    Sql = "CREATE TEMPORARY TABLE tmpscaalp ( " '
    Sql = Sql & "`numalbar` varchar(10) NOT NULL default '', "
    Sql = Sql & "`fechaalb` date NOT NULL default '0000-00-00', "
    Sql = Sql & "`codprove` int(6) unsigned NOT NULL default '0',"
    Sql = Sql & "`nomprove` varchar(40) NOT NULL, "
    Sql = Sql & "`domprove` varchar(35) NOT NULL, "
    Sql = Sql & "`codpobla` varchar(6) NOT NULL default '46',"
    Sql = Sql & "`pobprove` varchar(30) NOT NULL default 'A',"
    Sql = Sql & "`proprove` varchar(30) NOT NULL default 'A',"
    Sql = Sql & "`nifprove` varchar(15) NOT NULL default 'A',"
    Sql = Sql & "`telprove` varchar(15) default NULL,"
    Sql = Sql & "`codforpa` smallint(2) NOT NULL default '0',"
    Sql = Sql & "`dtoppago` decimal(4,2) NOT NULL default '0.00',"
    Sql = Sql & "`dtognral` decimal(4,2) NOT NULL default '0.00',"
    Sql = Sql & "`fecturno` date NOT NULL default '0000-00-00', "
    Sql = Sql & "`codturno` tinyint(1) NOT NULL) "
    
    Conn.Execute Sql
    
    'tabla temporal con la que cargaremos: slialp
    Sql = "CREATE TEMPORARY TABLE tmpslialp ( " 'TEMPORARY
    Sql = Sql & "`numalbar` varchar(10) NOT NULL default '',"
    Sql = Sql & "`fechaalb` date NOT NULL default '0000-00-00',"
    Sql = Sql & "`codprove` int(6) unsigned NOT NULL default '0',"
    Sql = Sql & "`numlinea` smallint(5) unsigned NOT NULL default '0',"
    Sql = Sql & "`codartic` int(6) NOT NULL,"
    Sql = Sql & "`codalmac` smallint(3) unsigned NOT NULL default '0',"
    Sql = Sql & "`nomartic` varchar(40) NOT NULL default '',"
    Sql = Sql & "`ampliaci` varchar(60) default NULL, "
    Sql = Sql & "`cantidad` decimal(12,2) default NULL,"
    Sql = Sql & "`precioar` decimal(10,5) NOT NULL default '0.00000',"
    Sql = Sql & "`dtoline1` decimal(4,2) NOT NULL default '0.00',"
    Sql = Sql & "`dtoline2` decimal(4,2) NOT NULL default '0.00',"
    Sql = Sql & "`importel` decimal(12,2) NOT NULL default '0.00',"
    Sql = Sql & "`fechahora` datetime)"
    
    Conn.Execute Sql
     
    CrearTMP = True
    
ECrear:
     If Err.Number <> 0 Then
        CrearTMP = False
        'Borrar la tabla temporal
        Sql = " DROP TABLE IF EXISTS tmpscaalp;"
        Conn.Execute Sql
        Sql = " DROP TABLE IF EXISTS tmpslialp;"
        Conn.Execute Sql
    End If
End Function


Private Sub BorrarTMP()
On Error Resume Next

    Conn.Execute " DROP TABLE IF EXISTS tmpslialp;"
    Conn.Execute " DROP TABLE IF EXISTS tmpscaalp;"
    
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Function PasarTemporales() As Boolean
Dim Sql As String
Dim RS As ADODB.Recordset

On Error GoTo ePasar

    Conn.Execute "delete from tmpinformes where codusu = " & vSesion.Codigo
    
    ' insertamos en tmpinformes: los albaranes que ya estaban en la scaalp CAMPO1 = 1
    Sql = "insert into tmpinformes (codusu, nombre1, fecha1, codigo1, campo1) "
    Sql = Sql & " select " & vSesion.Codigo & ", numalbar, fechaalb, codprove, 1 from tmpscaalp "
    Sql = Sql & " where (numalbar, fechaalb, codprove) in (select numalbar,fechaalb,codprove from scaalp) "

    Conn.Execute Sql


    Conn.Execute " INSERT INTO scaalp (numalbar,fechaalb,codprove,nomprove,domprove,codpobla,pobprove,proprove,nifprove,telprove,codforpa,dtoppago,dtognral,fecturno,codturno) SELECT * FROM tmpscaalp where (numalbar, fechaalb, codprove) not in (select nombre1,fecha1,codigo1 from tmpinformes where codusu = " & vSesion.Codigo & ") ; "
    Conn.Execute " INSERT INTO slialp (numalbar,fechaalb,codprove,numlinea,codartic,codalmac,nomartic,ampliaci,cantidad,precioar,dtoline1,dtoline2,importel) SELECT numalbar,fechaalb,codprove,numlinea,codartic,codalmac,nomartic,ampliaci,cantidad,precioar,dtoline1,dtoline2,importel FROM tmpslialp where (numalbar, fechaalb, codprove) not in (select nombre1,fecha1,codigo1 from tmpinformes where codusu = " & vSesion.Codigo & ") ; "
    
    'aqui es donde tenemos que actualizar la cantidad en stock, la fecha y ultimo precio de compra del articulo
    Sql = "SELECT * FROM tmpslialp where (numalbar, fechaalb, codprove) not in (select nombre1,fecha1,codigo1 from tmpinformes where codusu = " & vSesion.Codigo & ")"
    Set RS = New ADODB.Recordset
    RS.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not RS.EOF
        Sql = "update sartic set ultpreci = " & DBSet(RS!precioar, "N") & _
              ", ultfecha = " & DBSet(txtCodigo(0).Text, "F") & _
              " where codartic = " & DBSet(RS!codartic, "N") & _
              " and ultfecha < " & DBSet(txtCodigo(0).Text, "F")
        Conn.Execute Sql
'        ' solo si tiene control de stock
'        If DevuelveValor("select ctrstock from sartic where codartic = " & DBSet(RS!codArtic, "N")) = 1 Then
            Sql = "update sartic set canstock = canstock + " & DBSet(RS!cantidad, "N") & _
                  " where codartic = " & DBSet(RS!codartic, "N")
            Conn.Execute Sql
'        End If
        ' falta insertar en la smoval
        Sql = "insert into smoval (codartic,codalmac,fechamov,horamovi,tipomovi,detamovi,cantidad,impormov,codigope,letraser,document,numlinea) values ("
        Sql = Sql & DBSet(RS!codartic, "N") & ",1,"
        Sql = Sql & DBSet(RS!FechaAlb, "F") & ","
        Sql = Sql & DBSet(RS!fechahora, "FH") & ","
        Sql = Sql & "'S','ALC'," & DBSet(RS!cantidad, "N") & ","
        Sql = Sql & DBSet(RS!importel, "N") & ","
        Sql = Sql & DBSet(RS!CodProve, "N") & ","
        Sql = Sql & ValorNulo & ","
        Sql = Sql & DBSet(RS!numalbar, "T") & ","
        Sql = Sql & DBSet(RS!numlinea, "N") & ")"
        
        Conn.Execute Sql
        
        RS.MoveNext
    Wend
    Set RS = Nothing
    
    ' actualizamos la fecha de ultimo movimiento del proveedor
    Sql = "SELECT * FROM tmpscaalp where (numalbar, fechaalb, codprove) not in (select nombre1,fecha1,codigo1 from tmpinformes where codusu = " & vSesion.Codigo & ")"
    Set RS = New ADODB.Recordset
    RS.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not RS.EOF
        Sql = "update proveedor set fechamov = " & DBSet(txtCodigo(0).Text, "F") & _
              " where codprove = " & DBSet(RS!CodProve, "N") & _
              " and fechamov < " & DBSet(txtCodigo(0).Text, "F")
        Conn.Execute Sql
        RS.MoveNext
    Wend
    Set RS = Nothing
    
    ' insertamos en tmpinformes: los proveedores que estan introducidos automaticamente CAMPO1 = 2
    Sql = "insert into tmpinformes (codusu, nombre1, fecha1, codigo1, campo1, nombre2) "
    Sql = Sql & " select " & vSesion.Codigo & ", '' ," & ValorNulo & ", codprove, 2, nomprove from proveedor where domprove = 'AUTOMATICO'"
    
    Conn.Execute Sql
    
    ' insertamos en tmpinformes: los articulos que estan introducidos automaticamente CAMPO1 = 3
    Sql = "insert into tmpinformes (codusu, nombre1, fecha1, codigo1, campo1, nombre2) "
    Sql = Sql & " select " & vSesion.Codigo & ", '', " & ValorNulo & ", codartic, 3, nomartic from sartic where artnuevo = 1 "
        
    Conn.Execute Sql
    
    ' insertamos en tmpinformes: las familias que se han generado automaticamente CAMPO1 = 4
    Sql = "insert into tmpinformes (codusu, nombre1, fecha1, codigo1, campo1, nombre2) "
    Sql = Sql & " select " & vSesion.Codigo & ", '', " & ValorNulo & ", codfamia, 4, nomfamia from sfamia where nomfamia = 'AUTOMATICO'"
        
    Conn.Execute Sql
    
    
    PasarTemporales = True
    Exit Function
ePasar:
    PasarTemporales = False
End Function



Private Function ComprobarFechaAlbaran(nomFich As String) As Boolean
Dim NF As Long
Dim cad As String
Dim i As Integer
Dim longitud As Long
Dim RS As ADODB.Recordset
Dim Rs1 As ADODB.Recordset
Dim Numreg As Long
Dim Sql As String
Dim SQL1 As String
Dim total As Long
Dim v_cant As Currency
Dim v_impo As Currency
Dim v_prec As Currency
Dim b As Boolean

    On Error GoTo eComprobarFechaAlbaran
    
    ComprobarFechaAlbaran = False
    
    Sql = "delete from tmpinformes where codusu = " & vSesion.Codigo
    Conn.Execute Sql
    
    
    NF = FreeFile
    Open nomFich For Input As #NF ' & "\BV" & Format(CDate(txtcodigo(0).Text), "ddmmyy") & "." & Format(txtcodigo(1).Text, "000") For Input As #NF
    
    Line Input #NF, cad
    i = 0
    
    lblProgres(0).Caption = "Insertando en Tabla temporal: " & nomFich
    longitud = FileLen(nomFich)
    
    Pb1.visible = True
    Me.Pb1.Max = longitud
    Me.Refresh
    Me.Pb1.Value = 0
    ' PROCESO DEL FICHERO COMPRAS

    b = True

    While Not EOF(NF) And b
        i = i + 1
        
        Me.Pb1.Value = Me.Pb1.Value + Len(cad)
        lblProgres(1).Caption = "Linea " & i
        Me.Refresh
        
        b = ComprobarFecha(cad)
        
        Line Input #NF, cad
    Wend
    Close #NF
    
    If cad <> "" Then
        i = i + 1
        
        Me.Pb1.Value = Me.Pb1.Value + Len(cad)
        lblProgres(1).Caption = "Linea " & i
        Me.Refresh
        
        b = ComprobarFecha(cad)
    End If
    
    Pb1.visible = False
    lblProgres(0).Caption = ""
    lblProgres(1).Caption = ""

    ComprobarFechaAlbaran = b
    Exit Function

eComprobarFechaAlbaran:
    ComprobarFechaAlbaran = False
End Function




Private Function ComprobarFecha(cad As String) As Boolean
Dim Sql As String

Dim Albaran As String
Dim fechahora As String

Dim Fecha As String
Dim Hora As String

Dim Mens As String


Dim codsoc As String

    On Error GoTo eComprobarFecha

    ComprobarFecha = True

    Albaran = Mid(cad, 92, 15)
    fechahora = Mid(cad, 122, 14)
    
    Fecha = Mid(fechahora, 7, 2) & "/" & Mid(fechahora, 5, 2) & "/" & Mid(fechahora, 1, 4)
    Hora = Mid(fechahora, 9, 2) & ":" & Mid(fechahora, 11, 2) & ":" & Mid(fechahora, 13, 2)

    
    'Comprobamos fechas
    If Not EsFechaOK(Fecha) Then
            Mens = "Fecha incorrecta"
            Sql = "insert into tmpinformes (codusu, importe1, fecha1, nombre1) values (" & _
                  vSesion.Codigo & "," & DBSet(Albaran, "T") & "," & DBSet(Fecha, "F") & "," & DBSet(Mens, "T") & ")"
            
            Conn.Execute Sql
    Else
        If CDate(Fecha) <> CDate(txtCodigo(0).Text) Then
            Mens = "Fecha incorrecta"
            Sql = "insert into tmpinformes (codusu, importe1, fecha1, nombre1) values (" & _
                  vSesion.Codigo & "," & DBSet(Albaran, "T") & "," & DBSet(Fecha, "F") & "," & DBSet(Mens, "T") & ")"
            
            Conn.Execute Sql
        End If
    End If
    
eComprobarFecha:
    If Err.Number <> 0 Then
        ComprobarFecha = False
    End If
End Function


' fichero de comprobacion
Private Function ProcesarFicheroRegaixo2() As Boolean
Dim NF As Long
Dim cad As String
Dim i As Integer
Dim longitud As Long
Dim RS As ADODB.Recordset
Dim Rs1 As ADODB.Recordset
Dim Numreg As Long
Dim Sql As String
Dim SQL1 As String
Dim total As Long
Dim v_cant As Currency
Dim v_impo As Currency
Dim v_prec As Currency
Dim b As Boolean

    On Error GoTo eProcesarFichero2
    
    ProcesarFicheroRegaixo2 = False
    
    Sql = "select * from tmptraspaso where codusu = " & vSesion.Codigo & " and cast(mid(fecha,1,8) as date) = " & DBSet(txtCodigo(0).Text, "F")
    Sql = Sql & " order by albaran "
    
    
    i = 0
    
    lblProgres(0).Caption = "Insertando en Tabla temporal: "
    longitud = TotalRegistrosConsulta(Sql)
    
    Pb1.visible = True
    Me.Pb1.Max = longitud
    Me.Refresh
    Me.Pb1.Value = 0
    ' PROCESO DEL FICHERO VENTAS.TXT

    b = True

    Set RS = New ADODB.Recordset
    RS.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    While Not RS.EOF And b
        i = i + 1
        
        Me.Pb1.Value = Me.Pb1.Value + 1
        lblProgres(1).Caption = "Linea " & i
        Me.Refresh
        b = ComprobarRegistroReg(RS)
        ' If Not b Then Stop
        RS.MoveNext
    Wend
    
    Pb1.visible = False
    lblProgres(0).Caption = ""
    lblProgres(1).Caption = ""

    ProcesarFicheroRegaixo2 = b
    Exit Function

eProcesarFichero2:
    ProcesarFicheroRegaixo2 = False
End Function

'fichero de proceso
Private Function ProcesarFicheroRegaixo() As Boolean
Dim NF As Long
Dim cad As String
Dim i As Integer
Dim longitud As Long
Dim RS As ADODB.Recordset
Dim Rs1 As ADODB.Recordset
Dim Numreg As Long
Dim Sql As String
Dim SQL1 As String
Dim total As Long
Dim v_cant As Currency
Dim v_impo As Currency
Dim v_prec As Currency
Dim b As Boolean
Dim NomFic As String

    On Error GoTo eProcesarFicheroRegaixo

    ProcesarFicheroRegaixo = False
    NF = FreeFile
    
    i = 0
    
    Sql = "select turno,albaran,factura,fecha,cliente,nomclien,tarjeta,matricula,km,producto,nomprodu,surtidor,manguera,"
    Sql = Sql & " nsuministro,precio,descuento,descuentoporc,iva,cantidad,idtipopago,desctipopago,nif,importe "
    Sql = Sql & " from tmptraspaso where codusu = " & vSesion.Codigo & " and not idtipopago in (select forpaalvic from sforpa where tipovale in (1,2)) "
    Sql = Sql & " and cast(mid(fecha,1,8) as date) = " & DBSet(txtCodigo(0).Text, "F")
    Sql = Sql & " order by albaran, turno, fecha "
    
    lblProgres(0).Caption = "Procesando Fichero "
    longitud = TotalRegistrosConsulta(Sql)
    
    Pb1.visible = True
    Me.Pb1.Max = longitud
    Me.Refresh
    Me.Pb1.Value = 0
        
    b = True
    
    Set RS = New ADODB.Recordset
    RS.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    
    While Not RS.EOF
        i = i + 1
        
        Me.Pb1.Value = Me.Pb1.Value + 1
        lblProgres(1).Caption = "Linea " & i
        Me.Refresh
        
        b = InsertarLineaReg(RS)
         
        If b Then b = InsertarLineaTurnoReg(RS)
        
        If b = False Then
            ProcesarFicheroRegaixo = False
            Exit Function
        End If
        
        RS.MoveNext
    Wend
    Set RS = Nothing
    
    If b Then b = InsertarRecaudacionReg()
    
    ProcesarFicheroRegaixo = b
    
    Pb1.visible = False
    lblProgres(0).Caption = ""
    lblProgres(1).Caption = ""

eProcesarFicheroRegaixo:


End Function

Private Function InsertarLineaReg(ByRef RS As ADODB.Recordset) As Boolean
Dim NumLin As String
Dim codpro As String
Dim articulo As String
Dim Familia As String
Dim Precio As String
Dim ImpDes As String
Dim CodIVA As String
Dim b As Boolean
Dim Codclave As String
Dim Sql As String

Dim Import As Currency

Dim Base As String
Dim NombreBase As String
Dim Turno As String
Dim NumAlbaran As String
Dim NumFactura As String
Dim IdVendedor As String
Dim NombreVendedor As String
Dim fechahora As String
Dim Fecha As String
Dim Hora As String
Dim CodigoCliente As String
Dim NombreCliente As String
Dim Matricula As String
Dim Tarjeta As String
Dim CodigoProducto As String
Dim Surtidor As String
Dim Manguera As String
Dim PrecioLitro As String
Dim Descuento As String
Dim PorcDescuento As Currency
Dim cantidad As String
Dim Importe As String
Dim IdTipoPago As String
Dim DescrTipoPago As String
Dim CodigoTipoPago As String
Dim NifCliente As String
Dim IdProducto As String

Dim c_Cantidad As Currency
Dim c_Importe As Currency
Dim c_Precio As Currency
Dim c_Descuento As Currency
Dim c_Vale As Currency
Dim c_Devolucion As Currency
Dim Tarje As String

Dim SqlVale As String
Dim RsVale As ADODB.Recordset


Dim Mens As String
Dim numlinea As Long

Dim codsoc As String
Dim forpa As String

Dim Kilometros As String
Dim NomArtic As String

    On Error GoTo eInsertarLinea

    InsertarLineaReg = True
    
    Turno = DBLet(RS!Turno, "N")
    
    NumAlbaran = DBLet(RS!Albaran, "N")
    NumFactura = DBLet(RS!Factura, "T")
    If NumFactura <> "" Then
        NumFactura = Mid(NumFactura, 5, Len(NumFactura) - 4)
    End If
    
    fechahora = DBLet(RS!Fecha, "T")
    Fecha = Mid(fechahora, 7, 2) & "/" & Mid(fechahora, 5, 2) & "/" & Mid(fechahora, 1, 4)
    Hora = Mid(fechahora, 9, 2) & ":" & Mid(fechahora, 11, 2) & ":" & Mid(fechahora, 13, 2)
    CodigoCliente = DBLet(RS!CLIENTE, "T")
    NombreCliente = DBLet(RS!nomclien, "T")
    
    Tarjeta = DBLet(RS!Tarjeta, "N")
    Matricula = DBLet(RS!Matricula, "T")
    IdProducto = DBLet(RS!PRODUCTO, "N")
    Surtidor = DBLet(RS!Surtidor, "N")
    Manguera = DBLet(RS!Manguera, "N")
    
    PrecioLitro = DBLet(RS!Precio, "N")
    cantidad = DBLet(RS!cantidad, "N")
    Importe = DBLet(RS!Importe, "N")
    IdTipoPago = DBLet(RS!IdTipoPago, "N")
    DescrTipoPago = DBLet(RS!desctipopago, "T")
    CodigoTipoPago = DBLet(RS!IdTipoPago, "N")
    NifCliente = DBLet(RS!NIF, "T")
    
    ' en caso de que el codigo de cliente y el nombre no me vengan cojo el asociado a la forma de pago
    If CodigoCliente = "" And NombreCliente = "" Then
        CodigoCliente = DevuelveDesdeBDNew(cPTours, "sforpa", "codsocio", "forpaalvic", IdTipoPago, "N")
        NombreCliente = DevuelveDesdeBDNew(cPTours, "ssocio", "nomsocio", "codsocio", CodigoCliente, "N")
        Tarjeta = CodigoCliente
    End If
    
    Kilometros = DBLet(RS!km, "N")
    PorcDescuento = DBLet(RS!descuentoporc, "N")
    Descuento = Round(PrecioLitro * PorcDescuento / 100, 3)
    
    If Trim(Importe) = "" Then
        Exit Function
    Else
        If CCur(Importe) = 0 Then Exit Function
    End If
    
    c_Cantidad = cantidad 'Round2(CCur(cantidad) / 100, 2)
    c_Importe = Importe 'Round2(CCur(Importe) / 100, 2)
    c_Precio = PrecioLitro - Descuento 'Round2(CCur(PrecioLitro) / 100000, 5)
    c_Descuento = Descuento 'Round2(CCur(Descuento) / 100000, 5)
    
    c_Vale = 0
    
    SqlVale = "select * from tmptraspaso where codusu = " & DBSet(vSesion.Codigo, "N") & " and albaran = " & DBSet(NumAlbaran, "N")
    SqlVale = SqlVale & " and idtipopago in (select forpaalvic from sforpa where tipovale = 1) "
    Set RsVale = New ADODB.Recordset
    RsVale.Open SqlVale, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not RsVale.EOF Then
        c_Vale = DBLet(RsVale!Importe, "N")
    End If
    Set RsVale = Nothing
    
    c_Importe = c_Importe + c_Vale
    
    ' lo mismo con la devolucion de billetes
    c_Devolucion = 0
    
    SqlVale = "select * from tmptraspaso where codusu = " & DBSet(vSesion.Codigo, "N") & " and albaran = " & DBSet(NumAlbaran, "N")
    SqlVale = SqlVale & " and idtipopago in (select forpaalvic from sforpa where tipovale = 2) "
    Set RsVale = New ADODB.Recordset
    RsVale.Open SqlVale, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not RsVale.EOF Then
        c_Devolucion = DBLet(RsVale!Importe, "N")
    End If
    Set RsVale = Nothing
    
    c_Importe = c_Importe + c_Devolucion
    
    
'    '### [Monica] 17/09/2007
'    'no insertamos aquellas lineas de albaran de importe = 0
'    Importe = DBSet(c_Importe, "N")
'    If Import = 0 Then
'        InsertarLinea = True
'        Exit Function
'    End If
'    'hasta aqui
    
    'VRS:4.0.1(0) actualizamos el precio de articulo
    Sql = "update sartic set preventa = " & DBSet(PrecioLitro, "N") & _
          " where codartic = " & DBSet(IdProducto, "N")
    Conn.Execute Sql
    
    If DevuelveValor("select ctrstock from sartic where codartic = " & DBSet(IdProducto, "N")) = 1 Then
        Sql = "update sartic set " & _
              "  canstock = canstock - " & DBSet(c_Cantidad, "N") & _
              " where codartic = " & DBSet(IdProducto, "N")
        Conn.Execute Sql
    End If
    
    ' insertamos en la tabla de albaranes
    Codclave = SugerirCodigoSiguienteStr("scaalb", "codclave")
    
    forpa = ""
    forpa = DevuelveDesdeBDNew(cPTours, "sforpa", "codforpa", "forpaalvic", IdTipoPago, "N")
    
    
    If Trim(NumFactura) <> "" Then
        codsoc = DevuelveDesdeBDNew(cPTours, "ssocio", "codsocio", "nifsocio", NifCliente, "T")
        
        '[Monica]04/01/2015: en el caso de venga una factura sin nif, cogemos el de la forma de pago
        If codsoc = "" Then
            CodigoCliente = DevuelveDesdeBDNew(cPTours, "sforpa", "codsocio", "forpaalvic", IdTipoPago, "N")
            NombreCliente = DevuelveDesdeBDNew(cPTours, "ssocio", "nomsocio", "codsocio", CodigoCliente, "N")
            Tarjeta = CodigoCliente
            If Tarjeta = "0" Then Tarjeta = CodigoCliente
        
        Else
            '[Monica]17/06/2013: miramos si la tarjeta viene con algun asterisco
            If Mid(Tarjeta, 1, 4) = "****" Or Trim(Tarjeta) = "0" Or InStr(1, Tarjeta, "*") <> 0 Then
                Tarjeta = codsoc
            Else '++monica: 15/02/2008 las tarjetas profesionales tienen 16 caracteres solo analizo los 8 últimos
                If Len(Trim(Tarjeta)) = 16 Then
                    Tarjeta = Mid(Tarjeta, 9, 16)
                End If
                '++
            End If
            'fechahora--> txtcodigo(0).Text & " " & Time
        End If
        
        
        Sql = "INSERT INTO scaalb (codclave, codsocio, numtarje, numalbar, fecalbar, horalbar, " & _
              "codturno, codartic, cantidad, preciove, importel, codforpa, matricul, codtraba, " & _
              "numfactu, numlinea, kilometros, dtoalvic, importevale) VALUES (" & DBSet(Codclave, "T") & "," & DBSet(codsoc, "N") & "," & DBSet(Tarjeta, "N") & "," & _
               DBSet(NumAlbaran, "T") & "," & DBSet(txtCodigo(0).Text, "F") & "," & DBSet(txtCodigo(0).Text & " " & Hora, "FH") & "," & DBSet(txtCodigo(1).Text, "N") & "," & _
               DBSet(IdProducto, "N") & "," & DBSet(c_Cantidad, "N") & "," & DBSet(c_Precio, "N") & "," & _
               DBSet(c_Importe, "N") & "," & DBSet(forpa, "N") & "," & DBSet(Matricula, "T") & "," & DBSet(IdVendedor, "N") & ","
    
        numlinea = SugerirCodigoSiguienteStr("scaalb", "numlinea", "numfactu = " & DBSet(NumFactura, "N"))
        Sql = Sql & DBSet(NumFactura, "N") & "," & DBSet(numlinea, "N") & ","
    Else
        If InStr(1, CodigoCliente, "1Z") <> 0 Then
            
            codsoc = DevuelveDesdeBDNew(cPTours, "ssocio", "codsocio", "nifsocio", NifCliente, "T")
            
            If Tarjeta = "0" Then
                Tarje = DevuelveDesdeBDNew(cPTours, "starje", "numtarje", "numtarje", Tarjeta, "T")
                If Tarje = "" Then Tarjeta = codsoc
            End If
            
            '[Monica]05/01/2015: si el socio es de catadau o llombai cogemos su forma de pago (la del cliente)
            Sql = "select codforpa from ssocio where codsocio = " & DBSet(codsoc, "N") & " and codcoope in (1,2) "
            If TotalRegistrosConsulta(Sql) <> 0 Then
                forpa = DevuelveValor(Sql)
            End If
            
            
            
            Sql = "INSERT INTO scaalb (codclave, codsocio, numtarje, numalbar, fecalbar, horalbar, " & _
                  "codturno, codartic, cantidad, preciove, importel, codforpa, matricul, codtraba, " & _
                  "numfactu, numlinea, kilometros, dtoalvic, importevale) VALUES (" & DBSet(Codclave, "T") & "," & DBSet(codsoc, "N") & "," & DBSet(Tarjeta, "N") & "," & _
                   DBSet(NumAlbaran, "T") & "," & DBSet(txtCodigo(0).Text, "F") & "," & DBSet(txtCodigo(0).Text & " " & Hora, "FH") & "," & DBSet(txtCodigo(1).Text, "N") & "," & _
                   DBSet(IdProducto, "N") & "," & DBSet(c_Cantidad, "N") & "," & DBSet(c_Precio, "N") & "," & _
                   DBSet(c_Importe, "N") & "," & DBSet(forpa, "N") & "," & DBSet(Matricula, "T") & "," & DBSet(IdVendedor, "N") & ","
            Sql = Sql & "0,0,"
        Else
        
            '[Monica]17/06/2013: miramos si la tarjeta viene con algun asterisco
'            If Mid(Tarjeta, 1, 4) = "****" Or Trim(Tarjeta) = "0" Or InStr(1, Tarjeta, "*") <> 0 Then
'                Tarjeta = CodigoCliente
'            Else '++monica: 15/02/2008 las tarjetas profesionales tienen 16 caracteres solo analizo los 8 últimos
'                If Len(Trim(Tarjeta)) = 16 Then
'                    Tarjeta = Mid(Tarjeta, 9, 16)
'                End If
'                '++
'            End If
            
            If Tarjeta = "0" Then
                'COGEMOS LA PRIMERA TARJETA DEPENDIENDO DEL TIPO DE ARTICULO
                Dim tipogaso As String
                tipogaso = DevuelveDesdeBD("tipogaso", "sartic", "codartic", IdProducto, "N")
                Select Case tipogaso
                    Case "3" ' bonificado
                        Tarjeta = DevuelveDesdeBDNew(cPTours, "starje", "numtarje", "tiptarje", "1", "N", , "codsocio", CodigoCliente, "N")
                    Case "0", "1", "2", "4"
                        Tarjeta = DevuelveValor("select numtarje from starje where tiptarje <> 1 and codsocio = " & DBSet(CodigoCliente, "N"))
                End Select
            End If
            
            '[Monica]05/01/2015: si el socio es de catadau o llombai cogemos su forma de pago (la del cliente)
            Sql = "select codforpa from ssocio where codsocio = " & DBSet(CodigoCliente, "N") & " and codcoope in (1,2) "
            If TotalRegistrosConsulta(Sql) <> 0 Then
                forpa = DevuelveValor(Sql)
            End If
            
            
            
            Sql = "INSERT INTO scaalb (codclave, codsocio, numtarje, numalbar, fecalbar, horalbar, " & _
                  "codturno, codartic, cantidad, preciove, importel, codforpa, matricul, codtraba, " & _
                  "numfactu, numlinea, kilometros, dtoalvic, importevale) VALUES (" & DBSet(Codclave, "T") & "," & DBSet(CodigoCliente, "N") & "," & DBSet(Tarjeta, "N") & "," & _
                   DBSet(NumAlbaran, "T") & "," & DBSet(txtCodigo(0).Text, "F") & "," & DBSet(txtCodigo(0).Text & " " & Hora, "FH") & "," & DBSet(txtCodigo(1).Text, "N") & "," & _
                   DBSet(IdProducto, "N") & "," & DBSet(c_Cantidad, "N") & "," & DBSet(c_Precio, "N") & "," & _
                   DBSet(c_Importe, "N") & "," & DBSet(forpa, "N") & "," & DBSet(Matricula, "T") & "," & DBSet(IdVendedor, "N") & ","
            Sql = Sql & "0,0,"
            
        End If
    End If
    
    '[monica]24/06/2013: añadimos los kilometros
    Sql = Sql & DBSet(Round2(ComprobarCero(Trim(Kilometros)) / 100, 0), "N", "S") & "," '& ")"
 
 
    '[Monica]24/08/2015: añadimos el descuento
    Sql = Sql & DBSet(c_Descuento, "N") & "," & DBSet(c_Vale, "N") & ")"
 
    Conn.Execute Sql
    
eInsertarLinea:
    If Err.Number <> 0 Then
        InsertarLineaReg = False
        MsgBox "Error en Insertar Linea " & Err.Description, vbExclamation
    End If
End Function

Private Function InsertarLineaTurnoReg(ByRef RS As ADODB.Recordset) As Boolean
Dim NF As Long
Dim i As Long
Dim longitud As Long


Dim codpro As String
Dim cantidad As String
Dim Precio As String
Dim Importe As String
Dim Sql As String
Dim NumLin As Long
Dim cWhere As String

Dim Surtidor As String
Dim Manguera As String
Dim Inicial As String
Dim Final As String
Dim vInicial As Currency
Dim vFinal As Currency

    On Error GoTo eInsertarLineaTurnoNew

    InsertarLineaTurnoReg = True

            
    codpro = DBLet(RS!PRODUCTO, "N")
    cantidad = DBLet(RS!cantidad, "N")
    Precio = DBLet(RS!Precio, "N")
    Importe = DBLet(RS!Importe, "N")
    Surtidor = DBLet(RS!Surtidor, "N")
    Manguera = DBLet(RS!Manguera, "N")
    
    Sql = ""
    Sql = DevuelveDesdeBDNew(cPTours, "sturno", "codturno", "fechatur", txtCodigo(0).Text, "F", , "codturno", txtCodigo(1).Text, "N", "codartic", codpro, "N")
    If Sql = "" Then
    
        cWhere = "fechatur=" & DBSet(txtCodigo(0).Text, "F") & " and codturno = " & DBSet(txtCodigo(1).Text, "N")
        NumLin = CLng(SugerirCodigoSiguienteStr("sturno", "numlinea", cWhere))
        'insertamos
        ' antes surtidor y manguera: 1,1,
        Sql = "INSERT INTO sturno (fechatur, codturno, numlinea, tiporegi, numtanqu, nummangu, " & _
              " codartic, litrosve, importel, containi, contafin, tipocred) VALUES (" & _
              DBSet(txtCodigo(0).Text, "F") & "," & DBSet(txtCodigo(1).Text, "N") & "," & DBSet(NumLin, "N") & ",2," & DBSet(Surtidor, "N") & "," & DBSet(Manguera, "N") & "," & _
              DBSet(codpro, "N") & "," & DBSet(cantidad, "N") & "," & DBSet(Importe, "N") & ",0,0,0)"
              
        Conn.Execute Sql
    Else
        'actualizamos
        Sql = "UPDATE sturno SET importel = importel + " & DBSet(Importe, "N") & ", litrosve = litrosve +  " & DBSet(cantidad, "N") & " WHERE fechatur = " & _
              DBSet(txtCodigo(0).Text, "F") & " AND codturno = " & DBSet(txtCodigo(1).Text, "N") & " AND codartic = " & _
              DBSet(codpro, "N")
              
        Conn.Execute Sql
    End If
            
eInsertarLineaTurnoNew:
    If Err.Number <> 0 Then
        InsertarLineaTurnoReg = False
        MsgBox "Error en Insertar Turno en " & Err.Description, vbExclamation
    End If
End Function

Private Function InsertarRecaudacionReg() As Boolean
Dim forpa As String
Dim Importe As String
Dim Sql As String
Dim vImporte As String
Dim vForpaVale As String
Dim IdTipoPago As String
Dim Existe As String

    On Error GoTo eInsertarRecaudacion

    InsertarRecaudacionReg = True
    
    Sql = "insert into srecau (fechatur, codturno, codforpa, importel, intconta) "
    Sql = Sql & " select " & DBSet(txtCodigo(0).Text, "F") & "," & DBSet(txtCodigo(1).Text, "N") & ", codforpa, sum(importel-coalesce(importevale,0)), 0 "
    Sql = Sql & " from scaalb where fecalbar = " & DBSet(txtCodigo(0).Text, "F") & " and codturno = " & DBSet(txtCodigo(1).Text, "N")
    Sql = Sql & " group by 1,2,3 "
    Sql = Sql & " order by 1,2,3 "
    
    Conn.Execute Sql

    Sql = "select sum(coalesce(importevale,0)) from scaalb where fecalbar = " & DBSet(txtCodigo(0).Text, "F") & " and codturno = " & DBSet(txtCodigo(1).Text, "N")
    vImporte = DevuelveValor(Sql)
    vForpaVale = DevuelveValor("select codforpa from sforpa where tipovale = 1")
    If vImporte <> 0 Then
        Sql = "insert into srecau (fechatur, codturno, codforpa, importel, intconta) values ("
        Sql = Sql & DBSet(txtCodigo(0).Text, "F") & "," & DBSet(txtCodigo(1).Text, "N") & "," & DBSet(vForpaVale, "N") & "," & DBSet(vImporte, "N") & ",0) "
    
        Conn.Execute Sql
    End If


eInsertarRecaudacion:
    If Err.Number <> 0 Then
        InsertarRecaudacionReg = False
        MsgBox "Error en Insertar Recaudacion en " & Err.Description, vbExclamation
    End If
    
End Function


