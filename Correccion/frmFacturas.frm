VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmFacturas 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Facturación por Cliente"
   ClientHeight    =   3105
   ClientLeft      =   45
   ClientTop       =   2430
   ClientWidth     =   6645
   Icon            =   "frmFacturas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3105
   ScaleWidth      =   6645
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
      Height          =   2955
      Left            =   30
      TabIndex        =   2
      Top             =   0
      Width           =   6375
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   4725
         TabIndex        =   1
         Top             =   2070
         Width           =   975
      End
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   3540
         TabIndex        =   0
         Top             =   2070
         Width           =   975
      End
   End
End
Attribute VB_Name = "frmFacturas"
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
Dim tabla As String
Dim Codigo As String 'Código para FormulaSelection de Crystal Report
Dim TipCod As String
Dim Orden1 As String 'Campo de Ordenacion (por codigo) para Cristal Report
Dim Orden2 As String 'Campo de Ordenacion (por nombre) para Cristal Report

Dim PrimeraVez As Boolean

Private Sub KEYpress(KeyAscii As Integer)
    If KeyAscii = 13 Then 'ENTER
        KeyAscii = 0
        SendKeys "{tab}"
    ElseIf KeyAscii = 27 Then Unload Me  'ESC
    End If
End Sub


Private Sub cmdAceptar_Click()
Dim cDesde As String, cHasta As String 'cadena codigo Desde/Hasta
Dim nDesde As String, nHasta As String 'cadena Descripcion Desde/Hasta
Dim cadTABLA As String, cOrden As String
Dim cadMen As String
Dim i As Byte
Dim SQL As String
Dim Sql2 As String
Dim Sql3 As String
Dim tipo As Byte
Dim nRegs As Integer
Dim NumError As Long

Dim Codigo As String
Dim baseimpo As Dictionary

Dim RS As ADODB.Recordset
Dim Rs1 As ADODB.Recordset

Dim SQL1 As String
Dim impuesto As Currency

    On Error GoTo eError


    Conn.BeginTrans

    SQL = "select * from schfac_borrame order by 1,2,3 "
    
    Set RS = New ADODB.Recordset
    RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    While Not RS.EOF
        
        Set baseimpo = New Dictionary
        ' cogemos el numero de factura de parametros
        
        Sql3 = "select slhfac_borrame.codartic, sartic.codigiva, sum(implinea) importel from slhfac_borrame, sartic "
        Sql3 = Sql3 & " where letraser = " & DBSet(RS!letraser, "T")
        Sql3 = Sql3 & " and numfactu = " & DBSet(RS!numfactu, "N")
        Sql3 = Sql3 & " and fecfactu = " & DBSet(RS!fecfactu, "F")
        Sql3 = Sql3 & " and slhfac_borrame.codartic = sartic.codartic "
        Sql3 = Sql3 & " group by 1, 2 "
        
        Set Rs1 = New ADODB.Recordset
        Rs1.Open Sql3, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    
        While Not Rs1.EOF
            ' tenemos que calcular el impuesto multiplicando cantidad de linea por impuesto por articulo
            Codigo = "codigiva"
            SQL1 = ""
            SQL1 = DevuelveDesdeBD("impuesto", "sartic", "codartic", DBLet(Rs1!codartic), "N", Codigo)
            If SQL1 = "" Then
                impuesto = 0
            Else
                impuesto = CCur(SQL1) ' Comprueba si es nulo y lo pone a 0 o ""
            End If
            
            baseimpo(Val(Codigo)) = DBLet(baseimpo(Val(Codigo)), "N") + DBLet(Rs1!importel, "N")
        
            Rs1.MoveNext
        Wend
                
        Set Rs1 = Nothing
                
        Dim Imptot(2)
        Dim Tipiva(2)
        Dim Impbas(2)
        Dim impiva(2)
        Dim PorIva(2)
        Dim TotFac
        
        TotFac = 0
        For i = 0 To 2
             Tipiva(i) = Null
             Imptot(i) = Null
             Impbas(i) = Null
             impiva(i) = Null
             PorIva(i) = Null
        Next i
        
        For i = 0 To baseimpo.Count - 1
            If RS!letraser = "I" Then
                Tipiva(0) = vParamAplic.TipoIvaExento
                If IsNull(Imptot(0)) Then Imptot(0) = 0
                Imptot(0) = Imptot(0) + baseimpo.Items(i)
                PorIva(0) = DevuelveDesdeBDNew(cConta, "tiposiva", "porceiva", "codigiva", CStr(Tipiva(0)), "N")
                Impbas(0) = Round2(Imptot(0) / (1 + (PorIva(0) / 100)), 2)
                impiva(0) = Imptot(0) - Impbas(0)
                TotFac = Imptot(0)
            Else
                If i <= 2 Then
                    Tipiva(i) = baseimpo.Keys(i)
                    Imptot(i) = baseimpo.Items(i)
                    PorIva(i) = DevuelveDesdeBDNew(cConta, "tiposiva", "porceiva", "codigiva", CStr(Tipiva(i)), "N")
                    Impbas(i) = Round2(Imptot(i) / (1 + (PorIva(i) / 100)), 2)
                    impiva(i) = Imptot(i) - Impbas(i)
                    TotFac = TotFac + Imptot(i)
                End If
            End If
        Next i
        
        Sql2 = "update schfac_borrame set totalfac = " & DBSet(TotFac, "N") & _
               ", baseimp1 = " & DBSet(Impbas(0), "N") & _
               ", baseimp2 = " & DBSet(Impbas(1), "N", "S") & _
               ", baseimp3 = " & DBSet(Impbas(2), "N", "S") & _
               ", impoiva1 = " & DBSet(impiva(0), "N") & _
               ", impoiva2 = " & DBSet(impiva(1), "N", "S") & _
               ", impoiva3 = " & DBSet(impiva(2), "N", "S") & _
               ", tipoiva1 = " & DBSet(Tipiva(0), "N") & _
               ", tipoiva2 = " & DBSet(Tipiva(1), "N", "S") & _
               ", tipoiva3 = " & DBSet(Tipiva(2), "N", "S") & _
               ", porciva1 = " & DBSet(PorIva(0), "N") & _
               ", porciva2 = " & DBSet(PorIva(1), "N", "S") & _
               ", porciva3 = " & DBSet(PorIva(2), "N", "S") & _
               " where letraser = " & DBSet(RS!letraser, "T") & " and numfactu = " & DBSet(RS!numfactu, "N") & " and fecfactu = " & DBSet(RS!fecfactu, "F")
            
        Conn.Execute Sql2
        
        Set baseimpo = Nothing
        RS.MoveNext
        
    Wend
    
    Set RS = Nothing

    Conn.CommitTrans
    MsgBox "Proceso realizado correctamente"
    Exit Sub

eError:
    Conn.RollbackTrans
    MsgBox "No se ha realizado el proceso"

End Sub
 

Private Sub cmdCancel_Click()
    Unload Me
End Sub


Private Sub Form_Load()
Dim h As Integer, w As Integer
Dim List As Collection

    tabla = "schfac_borrame"
    
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

