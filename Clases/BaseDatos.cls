VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BaseDatos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Private mtipo As String
Private Conn As adodb.Connection
Private mRs As adodb.Recordset

Public Property Let tipo(ByVal vData As String)
    mtipo = vData
End Property

Public Property Get tipo() As String
    tipo = mtipo
End Property

Public Property Let Con(ByVal vData As adodb.Connection)
    Set Conn = vData
End Property

Public Property Get Con() As adodb.Connection
    Set Con = Conn
End Property


Public Function abrir(vODBC As String, vUsu As String, vPass As String) As Boolean
    '-- Abrimos la conexi�n en funci�n del ODBC pasado
    Set Conn = New adodb.Connection

'    Conn.ConnectionString = "Provider=MSDASQL.1;Persist Security Info=False;" & _
'                            "Data Source=" & vODBC & _
'                            ";UID=" & vUsu & _
'                            ";PWD=" & vPass
    Conn.ConnectionString = "DSN=arigasol;DESC=MySQL ODBC 3.51 Driver DSN;DESC=;DATABASE=" & vODBC & ";UID=" & vUsu & ";PASSWORD=" & vPass & ";PORT=3306;OPTION=3;STMT=;" & _
                        ";Persist Security Info=true"

    Conn.Open
    abrir = True
End Function

Public Function ejecutar(vSQL As String) As Long
' ### [Monica] 10/10/2006 a�ado la rutina de errores
On Error GoTo eEjecutar

    ejecutar = 0
    If Not (Conn Is Nothing) Then
        Conn.Execute vSQL
    Else
        '-- Falta la conexi�n y no podemos continuar
    End If

' ### [Monica] 10/10/2006 a�ado la rutina de errores
eEjecutar:
    If Err.Number <> 0 Then
        ejecutar = Err.Number
'        MuestraError Err.Number, Err.Description
    End If

End Function

Public Function ejecutar2(vSQL As String, ByRef menserr As String) As Long
' ### [Monica] 10/10/2006 a�ado la rutina de errores
On Error GoTo eEjecutar

    ejecutar2 = 0
    If Not (Conn Is Nothing) Then
        Conn.Execute vSQL
    Else
        '-- Falta la conexi�n y no podemos continuar
    End If

' ### [Monica] 10/10/2006 a�ado la rutina de errores
eEjecutar:
    If Err.Number <> 0 Then
        ejecutar2 = Err.Number
        MuestraError Err.Number, menserr & " " & Err.Description
    End If

End Function

Public Function cursor(vSQL As String) As adodb.Recordset
    '-- Devuelve un recordset con los valores obtenidos por SQL
    Set mRs = New adodb.Recordset
    mRs.Open vSQL, Conn, adOpenForwardOnly
    Set cursor = mRs
End Function

Public Function Fecha(vData) As String
    '-- Devuelve un string con la fecha correspondiente
    If vData = "0:00:00" Then
        Fecha = "NULL"
    Else
        If Not IsNull(vData) And IsDate(vData) Then
            Select Case mtipo
                Case "ACCESS"
                    Fecha = "#" & Format(vData, "yyyy/mm/dd") & "#"
                Case "MYSQL"
                    Fecha = "'" & Format(vData, "yyyy-mm-dd") & "'"
                Case "MULTIBASE"
                    Fecha = "'" & Format(vData, "dd/mm/yyyy") & "'"
                Case Default
                    Fecha = "'" & Format(vData, "yyyy-mm-dd") & "'"
            End Select
        Else
            Fecha = "NULL"
        End If
    End If
End Function

Public Function Hora(vData) As String
    '-- Devuelve un string con la fecha correspondiente
    If vData = "0:00:00" Then
        Hora = "NULL"
    Else
        If Not IsNull(vData) And IsDate(vData) Then
            Select Case mtipo
                Case "ACCESS"
                    Hora = "#" & Format(vData, "hh:mm:ss") & "#"
                Case "MYSQL"
                    Hora = "'" & Format(vData, "hh:mm:ss") & "'"
                Case "MULTIBASE"
                    Hora = "'" & Format(vData, "hh:mm") & "'"
                Case Default
            End Select
        Else
            Hora = "NULL"
        End If
    End If
End Function

Public Function fechahora(vData) As String
    '-- Devuelve un string con la fecha correspondiente
    If vData = "00/00/0000 00:00:00" Then
        fechahora = "NULL"
    Else
        If Not IsNull(vData) And IsDate(vData) Then
            Select Case mtipo
                Case "ACCESS"
                    fechahora = "#" & Format(vData, "yyyy/mm/dd") & " " & Format(vData, "hh:mm:ss") & "#"
                Case "MYSQL"
                    fechahora = "'" & Format(vData, "yyyy-mm-dd") & " " & Format(vData, "hh:mm:ss") & "'"
                Case "MULTIBASE"
                    fechahora = "'" & Format(vData, "dd/mm/yyyy") & " " & Format(vData, "hh:mm:ss") & "'"
                Case Default
                    fechahora = "'" & Format(vData, "yyyy-mm-dd") & " " & Format(vData, "hh:mm:ss") & "'"
            End Select
        Else
            fechahora = "NULL"
        End If
    End If
End Function

Public Function numero(vData, Optional decimales As Integer) As String
    If Not IsNull(vData) Then
        numero = CStr(vData)
        If decimales Then
            'numero = TransformaComasPuntos(CStr(vData / (10 ^ decimales)))
            numero = TransformaComasPuntos(Format((vData / (10 ^ decimales)), "0." & String(decimales, "0")))
        Else
            numero = TransformaComasPuntos(CStr(vData))
        End If
    Else
        'antes
        'numero = "0"
        'ahora
        numero = ValorNulo
    End If
    If numero = "" Then numero = 0
End Function

Public Function Texto(vData) As String
    If Not IsNull(vData) Then
        Texto = "'" & vData & "'"
    Else
        Texto = "''"
    End If
End Function

Public Function Borrar() As String
    Select Case mtipo
        Case "ACCESS"
            Borrar = "delete * "
        Case "MYSQL"
            Borrar = "delete "
        Case "MULTIBASE"
            Borrar = "delete "
        Case Default
    End Select
End Function

' ### [Monica] 11/10/2006
' a�adida la parte de transacciones
Public Sub AbrirTrans()

    Conn.BeginTrans

End Sub

Public Sub RollbackTrans()

    Conn.RollbackTrans

End Sub

Public Sub CommitTrans()

    Conn.CommitTrans

End Sub



