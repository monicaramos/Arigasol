Attribute VB_Name = "ModExpViajeros"
Option Explicit

'Obtener el primer asiento libre para un expediente de grupos
'de un MINORISTA cuando la ventapax=GRUPO
Public Function ObtenerSigNumAsiento(numExped As String, codEmpre As String, numPax As Integer) As String
Dim aAsientos() As Byte 'si cada asiento esta ocupado o libre

', a_ocupats() As Variant, a_lliures() As Variant
Dim i As Integer ', j As Integer
Dim RS As ADODB.Recordset
Dim SQL As String
 
    On Error GoTo ErrAsien

    If numPax > 0 Then
        ReDim Preserve aAsientos(numPax - 1)

        'Inicializamos a 0, todos libres
        For i = 0 To numPax - 1
            aAsientos(i) = 0
        Next i
        
        'Ponemos los que estan ocupados
        SQL = "SELECT numasien FROM viagrupc "
        SQL = SQL & " WHERE numexped=" & numExped & " AND codempre=" & codEmpre
        Set RS = New ADODB.Recordset
        RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        While Not RS.EOF
            i = RS.Fields(0).Value
            aAsientos(i - 1) = 1
            RS.MoveNext
        Wend
        RS.Close
        Set RS = Nothing
        
        
        'cogemos el primero que este libre
        For i = 0 To numPax - 1
            If aAsientos(i) = 0 Then
                ObtenerSigNumAsiento = i + 1
                Exit For
            End If
        Next i
    End If
    
    Exit Function
    
ErrAsien:
    MuestraError Err.Number, "No se ha podido obtener un nº de asiento.", Err.Description
    
End Function

