Attribute VB_Name = "ModInformes"
Option Explicit


'==============================================================
'====== FUNCIONES GENERALES  PARA INFORMES ====================

'Esta funcion lo que hace es genera el valor del campo
'El campo lo coge del recordset, luego sera field(i), y el tipo es para añadirle
'las coimllas, o quitarlas comas
'  Si es numero viene un 1 si no nada
'## NO LA USO, UTILIZO DBSET
'Public Function ParaBD(ByRef campo As ADODB.Field, Optional EsNumerico As Byte) As String
'
'    If IsNull(campo) Then
'        ParaBD = "NULL"
'    Else
'        Select Case EsNumerico
'        Case 1
'            ParaBD = TransformaComasPuntos(CStr(campo))
'        Case 2
'            'Fechas
'            ParaBD = "'" & Format(CStr(campo), "dd/MM/yyyy") & "'"
'        Case Else
'            ParaBD = "'" & campo & "'"
'        End Select
'    End If
'    ParaBD = "," & ParaBD
'End Function


Public Sub AbrirListado(numero As Byte)
    Screen.MousePointer = vbHourglass
    frmListado.OpcionListado = numero
    frmListado.Show vbModal
    Screen.MousePointer = vbDefault
End Sub


Public Function AnyadirAFormula(ByRef cadFormula As String, arg As String) As Boolean
'Concatena los criterios del WHERE para pasarlos al Crystal como FormulaSelection
    If arg = "Error" Then
        AnyadirAFormula = False
        Exit Function
    ElseIf arg <> "" Then
        If cadFormula <> "" Then
            cadFormula = cadFormula & " AND " & arg
        Else
            cadFormula = arg
        End If
    End If
    AnyadirAFormula = True
End Function


Public Function RegistrosAListar(vSQL As String, Optional vBD As Byte) As Byte
'Devuelve si hay algun registro para mostrar en el Informe con la seleccion
'realizada. Si no hay nada que mostrar devuelve 0 y no abrirá el informe
Dim RS As ADODB.Recordset

    On Error Resume Next
    
    Set RS = New ADODB.Recordset
    If vBD = cConta Then
        RS.Open vSQL, ConnConta, adOpenForwardOnly, adLockPessimistic, adCmdText
    Else
        RS.Open vSQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    End If


'    Set RS = New ADODB.Recordset
'    RS.Open vSQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    RegistrosAListar = 0
    If Not RS.EOF Then
        If RS.Fields(0).Value > 0 Then RegistrosAListar = 1 'Solo es para saber que hay registros que mostrar
    End If
    RS.Close
    Set RS = Nothing

    If Err.Number <> 0 Then
        RegistrosAListar = 0
        Err.Clear
    End If
End Function




Public Function HayRegParaInforme(cTabla As String, cWhere As String) As Boolean
'Comprobar si hay registros a Mostrar antes de abrir el Informe
Dim SQL As String

    SQL = "Select count(*) FROM " & QuitarCaracterACadena(cTabla, "_1")
    If cWhere <> "" Then
        cWhere = QuitarCaracterACadena(cWhere, "{")
        cWhere = QuitarCaracterACadena(cWhere, "}")
        cWhere = QuitarCaracterACadena(cWhere, "_1")
        SQL = SQL & " WHERE " & cWhere
    End If
    
    If RegistrosAListar(SQL) = 0 Then
        MsgBox "No hay datos para mostrar en el Informe.", vbInformation
        HayRegParaInforme = False
    Else
        HayRegParaInforme = True
    End If
End Function



Public Function CadenaDesdeHasta(cadDesde As String, cadHasta As String, campo As String, TipoCampo As String, Optional nomcampo As String) As String
'Devuelve la cadena de seleccion: " (campo >= cadDesde and campo<=cadHasta) "
'para Crystal Report
Dim cadAux As String

    If Trim(cadDesde) = "" And Trim(cadHasta) = "" Then
        'Campo Desde y Hasta no tiene valor
            cadAux = ""
    Else
        'Campo DESDE
        If cadDesde <> "" Then
            Select Case TipoCampo
                Case "N"
                    cadAux = campo & " >= " & Val(cadDesde)
                Case "T"
                    cadAux = campo & " >= """ & cadDesde & """"
                Case "F"
                    cadAux = campo & " >= Date(" & Year(cadDesde) & "," & Month(cadDesde) & "," & Day(cadDesde) & ")"
            End Select
        End If
        
        'Campo HASTA
        If cadHasta <> "" Then
            If cadAux <> "" Then 'Hay campo Desde y campo Hasta
                'Comprobar Desde <= Hasta
                Select Case TipoCampo
                    Case "N"
                        If CSng(cadDesde) > CSng(cadHasta) Then
                            MsgBox "El campo Desde debe ser menor que el campo Hasta", _
                            vbExclamation, "Error de campo"
                            cadAux = "Error"
                        Else
                            cadAux = cadAux & " and " & campo & " <= " & Val(cadHasta)
                        End If
                        
                    Case "T"
                        If cadDesde > cadHasta Then
                            MsgBox "El campo Desde debe ser menor que el campo Hasta", _
                            vbExclamation, "Error de campo"
                            cadAux = "Error"
                        Else
                            cadAux = cadAux & " and " & campo & " <= """ & cadHasta & """"
                        End If
                    
                    Case "F"
                        If CDate(cadDesde) > CDate(cadHasta) Then
                            MsgBox "El campo Desde debe ser menor que el campo Hasta", _
                            vbExclamation, "Error de campo"
                            cadAux = "Error"
                        Else
                            cadAux = cadAux & " and " & campo & " <= Date(" & Year(cadHasta) & "," & Month(cadHasta) & "," & Day(cadHasta) & ")"
                        End If
                End Select
            Else 'No hay campo Desde. Solo hay campo Hasta
                Select Case TipoCampo
                    Case "N"
                        cadAux = campo & " <= " & Val(cadHasta)
                    Case "T"
                        cadAux = campo & " <= """ & cadHasta & """"
                    Case "F"
                        cadAux = campo & " <= Date(" & Year(cadHasta) & "," & Month(cadHasta) & "," & Day(cadHasta) & ")"
                End Select
            End If
        End If
    End If
    If cadAux <> "" And cadAux <> "Error" Then cadAux = "(" & cadAux & ")"
    CadenaDesdeHasta = cadAux
End Function


Public Function CadenaDesdeHastaBD(cadDesde As String, cadHasta As String, campo As String, TipoCampo As String) As String
'Devuelve la cadena de seleccion: " (campo >= valor1 and campo<=valor2) "
'Para MySQL
Dim cadAux As String

    If Trim(cadDesde) = "" And Trim(cadHasta) = "" Then
        'Campo Desde y Hasta no tiene valor
            cadAux = ""
    Else
        'Campo DESDE
        If cadDesde <> "" Then
            Select Case TipoCampo
                Case "N"
                    cadAux = campo & " >= " & Val(cadDesde)
                Case "T"
                    cadAux = campo & " >= """ & cadDesde & """"
                Case "F"
                    cadAux = "(" & campo & " >= '" & Format(cadDesde, FormatoFecha) & "')"
            End Select
        End If
        
        'Campo HASTA
        If cadHasta <> "" Then
            If cadAux <> "" Then 'Hay campo Desde y campo Hasta
                'Comprobar Desde <= Hasta
                Select Case TipoCampo
                    Case "N"
                        If CSng(cadDesde) > CSng(cadHasta) Then
                            MsgBox "El campo Desde debe ser menor que el campo Hasta", _
                            vbExclamation, "Error de campo"
                            cadAux = "Error"
                        Else
                            cadAux = cadAux & " and " & campo & " <= " & Val(cadHasta)
                        End If
                        
                    Case "T"
                        If CSng(cadDesde) > CSng(cadHasta) Then
                            MsgBox "El campo Desde debe ser menor que el campo Hasta", _
                            vbExclamation, "Error de campo"
                            cadAux = "Error"
                        Else
                            cadAux = cadAux & " and " & campo & " <= """ & cadHasta & """"
                        End If
                    
                    Case "F"
                        If CDate(cadDesde) > CDate(cadHasta) Then
                            MsgBox "El campo Desde debe ser menor que el campo Hasta", _
                            vbExclamation, "Error de campo"
                            cadAux = "Error"
                        Else
                            cadAux = cadAux & " and (" & campo & " <= '" & Format(cadHasta, FormatoFecha) & "')"
                        End If
                End Select
                
            Else 'No hay campo Desde. Solo hay campo Hasta
                Select Case TipoCampo
                    Case "N"
                        cadAux = campo & " <= " & Val(cadHasta)
                    Case "T"
                        cadAux = campo & " <= """ & cadHasta & """"
                    Case "F"
                        cadAux = campo & " <= '" & Format(cadHasta, FormatoFecha) & "'"
                End Select
            End If
        End If
    End If
    If cadAux <> "" And cadAux <> "Error" Then cadAux = "(" & cadAux & ")"
    CadenaDesdeHastaBD = cadAux
End Function


Public Function AnyadirParametroDH(param As String, codD As String, codH As String, nomD As String, nomH As String) As String
On Error Resume Next
    
    If codD <> "" Then
        param = param & "DESDE: " & codD
        If nomD <> "" Then param = param & " - " & nomD
    End If
    If codH <> "" Then
        param = param & "  HASTA: " & codH
        If nomH <> "" Then param = param & " - " & nomH
    End If
    
    AnyadirParametroDH = param & """|"
    If Err.Number <> 0 Then Err.Clear
End Function



Public Function QuitarCaracterACadena(cadForm As String, Caracter As String) As String
'IN: [cadForm] es la cadena en la que se eliminara todos los caractes iguales a la vble [Caracter]
'OUT: cadena sin los caracteres
'EJEMPLO: "{scaalb.numalbar}", "{"  -->  "scaalb.numalbar}"
Dim i As Integer
Dim J As Integer
Dim Aux As String

    Aux = cadForm
    i = InStr(1, Aux, Caracter, vbTextCompare)
    While i > 0
        i = InStr(1, Aux, Caracter, vbTextCompare)
        If i > 0 Then
            J = Len(Caracter)
            Aux = Mid(Aux, 1, i - 1) & Mid(Aux, i + J, Len(Aux) - 1)
        End If
    Wend
    QuitarCaracterACadena = Aux
End Function


'## Utilizo la funcion REPLACE
'Public Function SustituirCadenas(CADENA As String, cad1 As String, cad2 As String) As String
''IN: Cadena es la cadena de texto en la que se va a sustituir la cad1 por la cad2
''OUT: cadena con la sustitucion
'
''EJEMPLO: cadena = "scaalb.codtipom='ALV' AND scaalb.numalbar=1"
''         cad1 = "scaalb"
''         cad2 = "slialb"
'
''         Resultado = "slialb.codtipom='ALV' AND slialb.numalbar=1"
'
'Dim i As Integer
'Dim j As Integer
'Dim Aux As String
'
'    Aux = CADENA
'    Do
'        i = InStr(1, Aux, cad1, vbTextCompare)
'        If i > 0 Then
'            j = Len(cad1)
'            Aux = Mid(Aux, 1, i - 1) & cad2 & Mid(Aux, i + j, Len(Aux) - 1)
'        End If
'    Loop Until i <= 0
'    SustituirCadenas = Aux
'End Function




Public Function PonerParamRPT(indice As Byte, cadParam As String, numParam As Byte, nomDocu As String) As Boolean
Dim vParamRpt As CParamRpt 'Tipos de Documentos
Dim cad As String

    Set vParamRpt = New CParamRpt
    
    If vParamRpt.Leer(indice) = 1 Then
        cad = "No se han podido cargar los Parámetros de Tipos de Documentos." & vbCrLf
        MsgBox cad & "Debe configurar la aplicación.", vbExclamation
        Set vParamRpt = Nothing
        PonerParamRPT = False
        Exit Function
    Else
        If cadParam = "" Then
            cad = "|"
        Else
            cad = ""
        End If
        cad = cad & "pCodigoISO=""" & vParamRpt.CodigoISO & """|"
        If vParamRpt.CodigoRevision = -1 Then
            cad = cad & "pCodigoRev=""" & "" & """|"
        Else
            cad = cad & "pCodigoRev=""" & Format(vParamRpt.CodigoRevision, "00") & """|"
        End If
        numParam = numParam + 2
        If vParamRpt.LineaPie1 <> "" Then
            cad = cad & "pLinea1=""" & vParamRpt.LineaPie1 & """|"
            numParam = numParam + 1
        End If
        If vParamRpt.LineaPie2 <> "" Then
            cad = cad & "pLinea2=""" & vParamRpt.LineaPie2 & """|"
            numParam = numParam + 1
        End If
        If vParamRpt.LineaPie3 <> "" Then
            cad = cad & "pLinea3=""" & vParamRpt.LineaPie3 & """|"
            numParam = numParam + 1
        End If
        If vParamRpt.LineaPie4 <> "" Then
            cad = cad & "pLinea4=""" & vParamRpt.LineaPie4 & """|"
            numParam = numParam + 1
        End If
        If vParamRpt.LineaPie5 <> "" Then
            cad = cad & "pLinea5=""" & vParamRpt.LineaPie5 & """|"
            numParam = numParam + 1
        End If
        cadParam = cadParam & cad
        nomDocu = vParamRpt.Documento
        PonerParamRPT = True
        Set vParamRpt = Nothing
    End If
End Function



Public Sub PonerFrameVisible(ByRef vFrame As Frame, visible As Boolean, h As Integer, w As Integer)
'Pone el Frame Visible y Ajustado al Formulario, y visualiza los controles
    
        vFrame.visible = visible
        If visible = True Then
            'Ajustar Tamaño del Frame para ajustar tamaño de Formulario al del Frame
            vFrame.Top = -90
            vFrame.Left = 0
            vFrame.Width = w
            vFrame.Height = h
        End If
End Sub

Public Sub AbrirListadoOfer(numero As Integer)
'Abre el Form con los listados de Ofertas
    Screen.MousePointer = vbHourglass
    frmListadoOfer.OpcionListado = numero
    frmListadoOfer.Show vbModal
    Screen.MousePointer = vbDefault
End Sub


