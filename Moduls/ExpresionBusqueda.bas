Attribute VB_Name = "ExpresionBusqueda"

Public Function SeparaCampoBusqueda(Tipo As String, campo As String, Cadena As String, ByRef DevSQL As String) As Byte
Dim Cad As String
Dim Aux As String
Dim CH As String
Dim Fin As Boolean
Dim I, J As String

On Error GoTo ErrSepara
SeparaCampoBusqueda = 1
DevSQL = ""
Cad = ""
Select Case Tipo
Case "N"
    '----------------  NUMERICO  ---------------------
    I = CararacteresCorrectos(Cadena, "N")
    If I > 0 Then Exit Function  'Ha habido un error y salimos
    'Comprobamos si hay intervalo ':'
    I = InStr(1, Cadena, ":")
    If I > 0 Then
        'Intervalo numerico
        Cad = Mid(Cadena, 1, I - 1)
        Aux = Mid(Cadena, I + 1)
        If Not IsNumeric(Cad) Or Not IsNumeric(Aux) Then Exit Function  'No son numeros
        'Intervalo correcto
        'Construimos la cadena
        
        '+-+-+- 23/05/2005 Canvi de Cèsar: per a que DevSQL tinga este aspecte: (taula.camp) >= 5 AND (taula.camp) <= 7 -+-+-+-+-
        'DevSQL = campo & " >= " & cad & " AND " & campo & " <= " & Aux
        DevSQL = campo & " >= " & Cad & ") AND (" & campo & " <= " & Aux
        '+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-
        
        'ELSE
        Else
            'Prueba
            'Comprobamos que no es el mayor
            If Cadena = ">>" Or Cadena = "<<" Then
                DevSQL = "1=1"
             Else
                    Fin = False
                    I = 1
                    Cad = ""
                    Aux = "NO ES NUMERO"
                    While Not Fin
                        CH = Mid(Cadena, I, 1)
                        If CH = ">" Or CH = "<" Or CH = "=" Then
                            Cad = Cad & CH
                            Else
                                Aux = Mid(Cadena, I)
                                Fin = True
                        End If
                        I = I + 1
                        If I > Len(Cadena) Then Fin = True
                    Wend
                    'En aux debemos tener el numero
                    If Not IsNumeric(Aux) Then Exit Function
                    'Si que es numero. Entonces, si Cad="" entronces le ponemos =
                    If Cad = "" Then Cad = " = "
                    DevSQL = campo & " " & Cad & " " & Aux
            End If
        End If
Case "F"
     '---------------- FECHAS ------------------
    I = CararacteresCorrectos(Cadena, "F")
    If I = 1 Then Exit Function
    'Comprobamos si hay intervalo ':'
    I = InStr(1, Cadena, ":")
    If I > 0 Then
        'Intervalo de fechas
        Cad = Mid(Cadena, 1, I - 1)
        Aux = Mid(Cadena, I + 1)
        If Not EsFechaOKString(Cad) Or Not EsFechaOKString(Aux) Then Exit Function  'Fechas incorrectas
        'Intervalo correcto
        'Construimos la cadena
        Cad = Format(Cad, FormatoFecha)
        Aux = Format(Aux, FormatoFecha)
        'En my sql es la ' no el #
        'DevSQL = Campo & " >=#" & Cad & "# AND " & Campo & " <= #" & AUX & "#"
        DevSQL = campo & " >='" & Cad & "' AND " & campo & " <= '" & Aux & "'"
        '----
        'ELSE
        Else
            'Comprobamos que no es el mayor
            If Cadena = ">>" Or Cadena = "<<" Then
                  DevSQL = "1=1"
            Else
                Fin = False
                I = 1
                Cad = ""
                Aux = "NO ES FECHA"
                While Not Fin
                    CH = Mid(Cadena, I, 1)
                    If CH = ">" Or CH = "<" Or CH = "=" Then
                        Cad = Cad & CH
                        Else
                            Aux = Mid(Cadena, I)
                            Fin = True
                    End If
                    I = I + 1
                    If I > Len(Cadena) Then Fin = True
                Wend
                'En aux debemos tener el numero
                If Not EsFechaOKString(Aux) Then Exit Function
                'Si que es numero. Entonces, si Cad="" entronces le ponemos =
                If Not Left(campo, 1) = "{" Then
                    Aux = "'" & Format(Aux, FormatoFecha) & "'"
                Else
                    Aux = "Date(" & Year(Aux) & "," & Month(Aux) & "," & Day(Aux) & ")"
                End If
                If Cad = "" Then Cad = " = "
                DevSQL = campo & " " & Cad & " " & Aux
            End If
        End If
    
Case "FHF" 'monica
    I = CararacteresCorrectos(Cadena, "F")
    If I = 1 Then Exit Function
    'Comprobamos si hay intervalo ':'
    I = InStr(1, Cadena, ":")
    If I > 0 Then
        'Intervalo de fechas
        Cad = Mid(Cadena, 1, I - 1)
        Aux = Mid(Cadena, I + 1)
        If Not EsFechaOKString(Cad) Or Not EsFechaOKString(Aux) Then Exit Function  'Fechas incorrectas
        'Intervalo correcto
        'Construimos la cadena
        Cad = Format(Cad, FormatoFecha)
        Aux = Format(Aux, FormatoFecha)
        'En my sql es la ' no el #
        'DevSQL = Campo & " >=#" & Cad & "# AND " & Campo & " <= #" & AUX & "#"
        DevSQL = "date(" & campo & ") >='" & Cad & "' AND date(" & campo & ") <= '" & Aux & "'"
        '----
        'ELSE
        Else
            'Comprobamos que no es el mayor
            If Cadena = ">>" Or Cadena = "<<" Then
                  DevSQL = "1=1"
            Else
                Fin = False
                I = 1
                Cad = ""
                Aux = "NO ES FECHA"
                While Not Fin
                    CH = Mid(Cadena, I, 1)
                    If CH = ">" Or CH = "<" Or CH = "=" Then
                        Cad = Cad & CH
                        Else
                            Aux = Mid(Cadena, I)
                            Fin = True
                    End If
                    I = I + 1
                    If I > Len(Cadena) Then Fin = True
                Wend
                'En aux debemos tener el numero
                If Not EsFechaOKString(Aux) Then Exit Function
                'Si que es numero. Entonces, si Cad="" entronces le ponemos =
                If Not Left(campo, 1) = "{" Then
                    Aux = "'" & Format(Aux, FormatoFecha) & "'"
                Else
                    Aux = "Date(" & Year(Aux) & "," & Month(Aux) & "," & Day(Aux) & ")"
                End If
                If Cad = "" Then Cad = " = "
                DevSQL = "date(" & campo & ") " & Cad & " " & Aux
            End If
        End If

Case "FHH" 'monica
    I = CararacteresCorrectos(Cadena, "F")
    If I = 1 Then Exit Function
    If Cadena = ">>" Or Cadena = "<<" Then
        DevSQL = "1=1"
    Else
        If EsHoraOK(Cadena) Then
            DevSQL = "time(" & campo & ") ='" & Cadena & "'"
        End If
    End If
    
Case "T"
    '---------------- TEXTO ------------------
    I = CararacteresCorrectos(Cadena, "T")
    If I = 1 Then Exit Function
    
    'Comprobamos que no es el mayor
    If Cadena = ">>" Or Cadena = "<<" Then
        DevSQL = "1=1"
        Exit Function
    End If
    
    'Comprobamos si hay intervalo ':'
    I = InStr(1, Cadena, ":")
    If I > 0 Then
        'Intervalo numerico
        Cad = Mid(Cadena, 1, I - 1)
        Aux = Mid(Cadena, I + 1)
'        If Not IsNumeric(cad) Or Not IsNumeric(Aux) Then Exit Function  'No son numeros
        'Intervalo correcto
        'Construimos la cadena
        DevSQL = campo & " >= '" & Cad & "' AND " & campo & " <= '" & Aux & "'"
    Else
    
        'Comprobamos si es LIKE o NOT LIKE
        Cad = Mid(Cadena, 1, 2)
        If Cad = "<>" Then
            Cadena = Mid(Cadena, 3)
            
            '[Monica]06/05/2019: condicion para no poner asteriscos
            '[Monica]28/11/2018: añado condicion para que T1 no concatene asteriscos
            If Tipo = "T" Then
                '[Monica]16/01/2014: añadida esta condicion para que ya no sean necesarios los asteriscos
                If Len(Cadena) <> 0 Then
                    If InStr(1, Cadena, "*") = 0 Then
                        Cadena = "*" & Cadena & "*"
                    End If
                End If
                '        16/01/2014: hasta aqui
            End If
            
            
            
            If Left(campo, 1) <> "{" Then
                'No es consulta seleccion para Report.
                DevSQL = campo & " NOT LIKE '"
            Else
                'Consulta de seleccion para Crystal Report
                DevSQL = "NOT (" & campo & " LIKE """ & Cadena & """)"
            End If
        Else
            '[Monica]06/05/2019: condicion para no poner asteriscos
            '[Monica]28/11/2018: añado condicion para que T1 no concatene asteriscos
            If Tipo = "T" Then
                '[Monica]16/01/2014: añadida esta condicion para que ya no sean necesarios los asteriscos
                If Len(Cadena) <> 0 Then
                    If InStr(1, Cadena, "*") = 0 Then
                        Cadena = "*" & Cadena & "*"
                    End If
                End If
                '        16/01/2014: hasta aqui
            End If
            
            If Left(campo, 1) <> "{" Then
            'NO es para report
                DevSQL = campo & " LIKE '"
            Else  'Es para report
                I = InStr(1, Cadena, "*")
                'Poner Consulta de seleccion para Crystal Report
                If I > 0 Then
                    DevSQL = campo & " LIKE """ & Cadena & """"
                Else
                    DevSQL = campo & " = """ & Cadena & """"
                End If
            End If
        End If
        
    
        'Cambiamos el * por % puesto que en ADO es el caraacter para like
        I = 1
        Aux = Cadena
        If Not Left(campo, 1) = "{" Then
          'No es para report
           While I <> 0
               I = InStr(1, Aux, "*")
               If I > 0 Then
                    Aux = Mid(Aux, 1, I - 1) & "%" & Mid(Aux, I + 1)
                End If
            Wend
        End If
        
        'Cambiamos el ? por la _ pue es su omonimo
        I = 1
        While I <> 0
            I = InStr(1, Aux, "?")
            If I > 0 Then Aux = Mid(Aux, 1, I - 1) & "_" & Mid(Aux, I + 1)
        Wend
    
        'Poner el valor de la expresion
        If Left(campo, 1) <> "{" Then
            'No es consulta seleccion para Report.
            DevSQL = DevSQL & Aux & "'"
        'Else
            'Consulta de seleccion para Crystal Report
            'DevSQL = DevSQL & CADENA & """)"
        End If
    End If
    '=========
    'ANTES
'    If cad = "<>" Then
'        '====David
'        'Aux = Mid(CADENA, 3)
'        'LAura
'        Aux = Mid(Aux, 3)
'        '====
'        If Left(Campo, 1) <> "{" Then
'            'Mo es consulta seleccion para Report.
'            DevSQL = Campo & " NOT LIKE '" & Aux & "'"
'        Else
'            'Consulta de seleccion para Crystal Report
'            DevSQL = Campo & " <> " & Aux & ""
'        End If
'    Else
'        If Left(Campo, 1) <> "{" Then
'            DevSQL = Campo & " LIKE '" & Aux & "'"
'        ElseIf Left(Aux, 4) = "like" Then
'            'Consulta de seleccion para Crystal Report
'            DevSQL = Campo & " " & Aux
'        Else
'            'Consulta de seleccion para Crystal Report
'            DevSQL = Campo & " = """ & Aux & """"
'        End If
'    End If
    
    
Case "B"
    'Como vienen de check box o del option box
    'los escribimos nosotros luego siempre sera correcta la
    'sintaxis
    'Los booleanos. Valores buenos son
    'Verdadero , Falso, True, False, = , <>
    'Igual o distinto
    I = InStr(1, Cadena, "<>")
    If I = 0 Then
        'IGUAL A valor
        Cad = " = "
        Else
            'Distinto a valor
        Cad = " <> "
    End If
    'Verdadero o falso
    I = InStr(1, Cadena, "V")
    If I > 0 Then
            Aux = "True"
            Else
            Aux = "False"
    End If
    'Ponemos la cadena
    DevSQL = campo & " " & Cad & " " & Aux
    
Case Else
    'No hacemos nada
        Exit Function
End Select
SeparaCampoBusqueda = 0
ErrSepara:
    If Err.Number <> 0 Then MuestraError Err.Number
End Function


Public Function CararacteresCorrectos(vcad As String, Tipo As String) As Byte
Dim I As Integer
Dim CH As String
Dim Error As Boolean

CararacteresCorrectos = 1
Error = False
Select Case Tipo
Case "N"
    'Numero. Aceptamos numeros, >,< = :
    For I = 1 To Len(vcad)
        CH = Mid(vcad, I, 1)
        Select Case CH
            Case "0" To "9"
            Case "<", ">", ":", "=", ".", " ", "-"
            Case Else
                Error = True
                Exit For
        End Select
    Next I
Case "T"
    'Texto aceptamos numeros, letras y el interrogante y el asterisco
    For I = 1 To Len(vcad)
        CH = Mid(vcad, I, 1)
        Select Case CH
            Case "a" To "z"
            Case "è", "é", "í" 'Añade Laura: 16/03/06
            Case "A" To "Z"
            Case "0" To "9"
            Case "*", "%", "?", "_", "\", "/", ":", ".", " " ' estos son para un caracter sol no esta demostrado , "%", "&"
            'Esta es opcional
            Case "<", ">"
            Case "Ñ", "ñ"
            Case "-" 'Añade Laura
            Case "#", "@", "$" '[Monica]29/07/2016: no estaba incluida la arroba
            Case Else
                Error = True
                Exit For
        End Select
    Next I
    
Case "F"
    'Tipo Fecha. Aceptamos Numeros , "/" ,":"
    For I = 1 To Len(vcad)
        CH = Mid(vcad, I, 1)
        Select Case CH
            Case "0" To "9"
            Case "<", ">", ":", "/", "="
            Case Else
                Error = True
                Exit For
        End Select
    Next I

Case "B"
    'Numeros , "/" ,":"
    For I = 1 To Len(vcad)
        CH = Mid(vcad, I, 1)
        Select Case CH
            Case "0" To "9"
            Case "<", ">", ":", "/", "=", " "
            Case Else
                Error = True
                Exit For
        End Select
    Next I
End Select
'Si no ha habido error cambiamos el retorno
If Not Error Then CararacteresCorrectos = 0
End Function


Public Function QuitarCaracterEnter(vcad As String) As String
Dim I As Integer

    Do
        I = InStr(1, vcad, Chr(13))
        If I > 0 Then 'Hay ENTER
            vcad = Mid(vcad, 1, I - 1) & Mid(vcad, I + 2)
        End If
    Loop Until I = 0
    QuitarCaracterEnter = vcad
End Function



Public Function ContieneCaracterBusqueda(Cadena As String) As Boolean
'Comprueba si la cadena contiene algun caracter especial de busqueda
' >,>,>=,: , ....
'si encuentra algun caracter de busqueda devuelve TRUE y sale
Dim b As Boolean
Dim I As Integer

    'For i = 1 To Len(cadena)
    I = 1
    b = False
    Do
        CH = Mid(Cadena, I, 1)
        Select Case CH
            Case "<", ">", ":", "="
                b = True
            Case "*", "%", "?", "_", "\", ":" ', "."
                b = True
            Case Else
                b = False
        End Select
    'Next i
        I = I + 1
    Loop Until (b = True) Or (I > Len(Cadena))
    ContieneCaracterBusqueda = b
End Function

Public Function QuitarCaracterNULL(vcad As String) As String
Dim I As Integer

    Do
        I = InStr(1, vcad, vbNullChar)
        If I > 0 Then 'Hay null
            vcad = Mid(vcad, 1, I - 1) & Mid(vcad, I + 2)
        End If
    Loop Until I = 0
    QuitarCaracterNULL = vcad
End Function

