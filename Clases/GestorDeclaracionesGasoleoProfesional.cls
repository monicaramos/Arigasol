VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GestorDeclaracionesGasoleoProfesional"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Private sql As String
Private rs As ADODB.Recordset
'Para activar este evento, use RaiseEvent con la siguiente sintaxis:
'RaiseEvent procesando[(arg1, arg2, ... , argn)]
Public Event procesando(registro As Integer)


Public Function declaraGasoleoProfesional(conn As ADODB.Connection, Optional test As Boolean) As Boolean
    Dim mDeclaracion As Declaracion
    Dim mColSuministro As ColSuministro
    Dim mSuministro As Suministro
    Dim i As Integer
On Error GoTo err_declaraGasoleoProfesional
    conn.Execute "START TRANSACTION"
        sql = "select a.numtarje, a.fecalbar, a.horalbar, a.cantidad, b.matricul, c.nifsocio, d.porcbd" & _
                " from scaalb as a, starje as b, ssocio as c, sartic as d" & _
                " where a.numtarje in (select numtarje from starje where tiptarje = 2)" & _
                " and a.codartic in (select codartic from sartic where gp = 1)" & _
                " and b.numtarje = a.numtarje" & _
                " and c.codsocio = a.codsocio" & _
                " and d.codartic = a.codartic" & _
                " and a.declaradogp = 0"
        Set rs = New ADODB.Recordset
        rs.Open sql, conn, adOpenForwardOnly
        If Not rs.EOF Then
            i = 0
            Set mDeclaracion = New Declaracion
            rs.MoveFirst
            '-- el id se montará al guardar la Declaracion
            '-- el codEE se obtiene automáticamente
            If test Then mDeclaracion.test = "S" Else mDeclaracion.test = "N"
            Set mColSuministro = New ColSuministro
            While Not rs.EOF
                i = i + 1
                RaiseEvent procesando(i)
                Set mSuministro = New Suministro
                With mSuministro
                    '-- el id se montará al guardar la Declaracion
                    '-- el idmovcont se genera automáticamente
                    .fecha = rs!fecAlbar
                    .Hora = rs!horalbar
                    .CodPro = rs!porcbd
                    .lit = rs!cantidad
                    .NIF = rs!nifsocio
                    .matricula = rs!matricul
                End With
                mColSuministro.Add mSuministro
                rs.MoveNext
            Wend
            Set mDeclaracion.suministros = mColSuministro
            If Not mDeclaracion.guardar(conn) Then
                conn.Execute "ROLLBACK"
                Exit Function
            End If
        End If
        '-- Ya lo hemos pasado a la base de datos y lo actualizamos como declarado
        sql = "update scaalb set declaradogp = 1" & _
                " where numtarje in (select numtarje from starje where tiptarje = 2)" & _
                " and codartic in (select codartic from sartic where gp = 1)"
        conn.Execute sql
    conn.Execute "COMMIT"
    declaraGasoleoProfesional = True
    Exit Function
err_declaraGasoleoProfesional:
    conn.Execute "ROLLBACK"
    MsgBox Err.Description & " (" & Err.Number & ")", vbExclamation, "declaraGasoleoProfesional"
End Function

Public Function enviaDeclaracionesPendientes(conn As ADODB.Connection) As Boolean
    Dim mDeclaracion As Declaracion
    Dim i As Integer
On Error GoTo err_enviaDeclaracionesPendientes
    conn.Execute "START TRANSACTION"
        sql = "select * from gp_suministrv2ent where situacion = 0"
        Set rs = New ADODB.Recordset
        rs.Open sql, conn, adOpenForwardOnly
        If Not rs.EOF Then
            i = 0
            rs.MoveFirst
            While Not rs.EOF
                i = i + 1
                RaiseEvent procesando(i)
                Set mDeclaracion = New Declaracion
                If Not mDeclaracion.leer(rs!id, conn) Then
                    conn.Execute "ROLLBACK"
                    Exit Function
                End If
                If Not mDeclaracion.exportarXML Then
                    conn.Execute "ROLLBACK"
                    Exit Function
                End If
                rs.MoveNext
            Wend
        End If
        If Not enviaAEAT() Then
            conn.Execute "ROLLBACK"
            Exit Function
        End If
        sql = "update gp_suministrv2ent set situacion = 1 where situacion = 0"
        conn.Execute sql
        enviaDeclaracionesPendientes = True
    conn.Execute "COMMIT"
    arrancaUtilidad
    Exit Function
err_enviaDeclaracionesPendientes:
    conn.Execute "ROLLBACK"
    MsgBox Err.Description & " (" & Err.Number & ")", vbExclamation, "enviaDeclaracionesPendientes"
End Function

Private Function enviaAEAT() As Boolean
    Dim ms As String
    Dim fichero As String
On Error GoTo err_enviaAEAT
    '-- Envia los fichero XML a la Agencia tributaria
    '-- 1* Comprueba que el directorio con la utilidad de la agencia existe y la utilidad se encuentra en el
    If Dir(App.path & "\aeat\aduanetxml_3_0.jar") = "" Then
        ms = "La utilidad de la agencia tributaria (aduanetxml_3_0.jar) no se encuentra en el subdirectorio aeat de la aplicación." & vbCrLf & _
                "Descárgela de la web de la agencia (http://www.agenciatributaria.es) y cópiela a " & App.path & "\aeat"
        MsgBox ms, vbExclamation
        Exit Function
    End If
    '-- 2* Comprobamos que la bandeja de salida está disponible
    If Dir(App.path & "\aeat\aduanetxml\BandejaSalida", vbDirectory) = "" Then
        ms = "No se encuentra la bandeja de salida en la que depositar los mensajes." & vbCrLf & _
                "Ejecute " & App.path & "\aeat\aduanetxml_3_0.jar y los directorios se crearán automáticamente."
        MsgBox ms, vbExclamation
        Exit Function
    End If
    '-- 3* Copiar los ficheros a la bandeja de salida y borrarlos de origen
    fichero = Dir(App.path & "\*.xml")
    Do While fichero <> ""
        FileCopy App.path & "\" & fichero, App.path & "\aeat\aduanetxml\BandejaSalida\" & fichero
        fichero = Dir
    Loop
    If Dir(App.path & "\*.xml") <> "" Then Kill App.path & "\*.xml"
    '-- 4 * Ya está, ahora si quiere puede arrancar la utilidad
    enviaAEAT = True
    Exit Function
err_enviaAEAT:
    MsgBox Err.Description & " (" & Err.Number & ")", vbExclamation, "enviaDeclaracionesPendientes"
End Function

Private Sub arrancaUtilidad()
    Dim ms As String
    On Error GoTo err_arrancaUtilidad
    '-- Arranca la utilidad de la agencia tributaria
    ms = "El proceso de traspaso de declaraciones ha finalizado corretamente." & _
            "¿Desea arrancar la utilidad de traspaso de la Agencia Tributaria?"
    If MsgBox(ms, vbYesNo + vbQuestion) = vbYes Then
        Shell "javaw -jar " & App.path & "\aeat\aduanetxml_3_0.jar"
    End If
    Exit Sub
err_arrancaUtilidad:
    MsgBox "No es posible arrancar automáticamente la utilidad. Hágalo manualmente", vbExclamation
End Sub

Public Function quedaPorDeclarar(conn As ADODB.Connection, fecha As Date) As Boolean
    '-- Comprueba si por debajo de una fecha determinada quedan operaciones por declarar.
On Error GoTo err_quedaPorDeclarar
    sql = "select a.numtarje, a.fecalbar, a.horalbar, a.cantidad, b.matricul, c.nifsocio, d.porcbd" & _
            " from scaalb as a, starje as b, ssocio as c, sartic as d" & _
            " where a.numtarje in (select numtarje from starje where tiptarje = 2)" & _
            " and a.codartic in (select codartic from sartic where gp = 1)" & _
            " and b.numtarje = a.numtarje" & _
            " and c.codsocio = a.codsocio" & _
            " and d.codartic = a.codartic" & _
            " and a.declaradogp = 0" & _
            " and a.fecalbar <= '" & Format(fecha, "yyyy-mm-dd") & "'"
    Set rs = New ADODB.Recordset
    rs.Open sql, conn, adOpenForwardOnly
    If Not rs.EOF Then quedaPorDeclarar = True
    Exit Function
err_quedaPorDeclarar:
    MsgBox Err.Description & " (" & Err.Number & ")", vbExclamation, "enviaDeclaracionesPendientes"
End Function
