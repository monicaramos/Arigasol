Attribute VB_Name = "libIBAN"
Option Explicit




'A partir de una cuenta banco formateada y todos los numeros juntos (chr(20))
'  devuelve DOS(2) caracteres del IBAN
'  calculados como dice la formula
'  i=ctabanco_con ES... mod 97
'  i = 98-i
' format(i,"00"             'para que copie                     'es lo que devuelve
'
'Puede NO poner pais. Sera ES
Public Function DevuelveIBAN2(PAIS As String, ByVal CtaBancoFormateada As String, DosCaracteresIBAN As String) As Boolean
Dim Aux As String
Dim N As Long
Dim CadenaPais As String
On Error GoTo EDevuelveIBAN
    DevuelveIBAN2 = False
    DosCaracteresIBAN = ""
    
    
    
    If PAIS = "" Then
        PAIS = "ES"
    Else
        If Len(PAIS) <> 2 Then
            PAIS = "ES"
        Else
            PAIS = UCase(PAIS)
        End If
    End If
    
    
    'Ejemplo mio: 20770294901101867914  IBAN: 41
    'Construir el IBAn:
    'A la derecha de la cuenta se pone
    '   el ES00-->   20770294961101915202ES00 ->92
    'Se transforma las letras ES a numeros.
    ' E=14 S=28
    '           ->>> 20770294961101915202 142800
    If PAIS = "ES" Then
        CadenaPais = "1428"
    Else
        N = Asc(Mid(PAIS, 1, 1))
        If N < 65 Or N > 90 Then Err.Raise 513, , "LEtra incorrecta PAIS: " & PAIS
        N = N - 55
        CadenaPais = CStr(N)
        N = Asc(Mid(PAIS, 2, 1))
        If N < 65 Or N > 90 Then Err.Raise 513, , "LEtra incorrecta PAIS: " & PAIS
        N = N - 55
        CadenaPais = CadenaPais & CStr(N)
    End If
    'Se le a�aden 2 ceros al final
    CadenaPais = CadenaPais & "00"
    'Esta es la cadena para ES. SiCadenaPais  fuera otro pais es aqui donde hay que cambiar
    CtaBancoFormateada = CtaBancoFormateada & "142800"
    Aux = ""
    While CtaBancoFormateada <> ""
        If Len(CtaBancoFormateada) >= 6 Then
            Aux = Aux & Mid(CtaBancoFormateada, 1, 6)
            CtaBancoFormateada = Mid(CtaBancoFormateada, 7)
        Else
            Aux = Aux & CtaBancoFormateada
            CtaBancoFormateada = ""
        End If
        
        N = CLng(Aux)
        N = N Mod 97
        
        Aux = CStr(N)
    Wend
        
    N = 98 - N
    
    DosCaracteresIBAN = Format(N, "00")
    DevuelveIBAN2 = True
    Exit Function
EDevuelveIBAN:
    CadenaPais = Err.Description
    CadenaPais = Err.Number & "   " & CadenaPais
    MsgBox "Devuelve IBAN. " & vbCrLf & CadenaPais, vbExclamation
    Err.Clear
End Function




Public Function IBAN_Correcto(IBAN As String) As Boolean
Dim Aux As String
    IBAN_Correcto = False
    Aux = ""
    If Len(IBAN) <> 4 Then
        Aux = "Longitud incorrecta"
    Else
        If IsNumeric(Mid(Aux, 3, 2)) Then
            Aux = "Digitos 3 y 4 deben ser numericos"
        Else
            'Podriamos comprobar lista de paises
    
        End If
    End If
    If Aux <> "" Then
        MsgBox Aux, vbExclamation
    Else
        IBAN_Correcto = True
    End If
End Function



'A partir de una cadena, con letras y numeros convertira
'en mod 97,10 Norma ISO 7064
'Para ello los caracteres se pasan a dos digitos
Public Function CadenaTextoMod97(Cadena As String) As String
Dim i As Integer
Dim C As String
Dim N As Long

    Cadena = Trim(Cadena)
    C = ""
    'Substitucion de texto por caracteres
    For i = 1 To Len(Cadena)
        N = Asc(Mid(Cadena, i, 1))
        If N >= 48 Then
            If N <= 57 Then
                'Es numerico 0..9
                'C = C & CStr(N)
            Else
                If N < 65 Or N > 90 Then
                    'MAL. No es un caracter ASCII entre A..Z  (10..35)
                    N = 0
                Else
                    N = N - 55  'el ascci menos 55 (0...35)
                End If
            End If
        End If
        If N = 0 Then
            CadenaTextoMod97 = "Caracter NO valido: " & Mid(Cadena, i, 1) & " --- " & Cadena
            Exit Function
        Else
            If N >= 48 Then
                'Es un numero
                C = C & Chr(N)
            Else
                C = C & CStr(N)
            End If
        End If
        
    Next
    
    
    
    'Ya tengo C que es numerica
    Cadena = C
    C = ""
    While Cadena <> ""
        If Len(Cadena) >= 6 Then
            C = C & Mid(Cadena, 1, 6)
            Cadena = Mid(Cadena, 7)
        Else
            C = C & Cadena
            Cadena = ""
        End If
        
        N = CLng(C)
        N = N Mod 97
        
        C = CStr(N)
    Wend
        
    N = 98 - N
    CadenaTextoMod97 = Format(N, "00")
End Function


Public Function DigitoControlCorrecto(CC As String) As String
    Dim ent As String ' Entidad
    Dim Suc As String ' Oficina
    Dim DC As String ' Digitos de control
    Dim i, i2, i3, i4 As Integer
    Dim NumCC As String ' N�mero de cuenta propiamente dicho
    '-- Esta funci�n comprueba la correcci�n de un n�mero de cuenta pasado en CC
    If Len(CC) <> 20 Then Exit Function '-- Las cuentas deben contener 20 d�gitos en total
    
    
    '-- Calculamos el primer d�gito de control
    i = Val(Mid(CC, 1, 1)) * 4
    i = i + Val(Mid(CC, 2, 1)) * 8
    i = i + Val(Mid(CC, 3, 1)) * 5
    i = i + Val(Mid(CC, 4, 1)) * 10
    i = i + Val(Mid(CC, 5, 1)) * 9
    i = i + Val(Mid(CC, 6, 1)) * 7
    i = i + Val(Mid(CC, 7, 1)) * 3
    i = i + Val(Mid(CC, 8, 1)) * 6
    i2 = Int(i / 11)
    i3 = i - (i2 * 11)
    i4 = 11 - i3
    Select Case i4
        Case 11
            i4 = 0
        Case 10
            i4 = 1
    End Select
    DC = i4
    '-- Calculamos el segundo d�gito de control
    i = Val(Mid(CC, 11, 1)) * 1
    i = i + Val(Mid(CC, 12, 1)) * 2
    i = i + Val(Mid(CC, 13, 1)) * 4
    i = i + Val(Mid(CC, 14, 1)) * 8
    i = i + Val(Mid(CC, 15, 1)) * 5
    i = i + Val(Mid(CC, 16, 1)) * 10
    i = i + Val(Mid(CC, 17, 1)) * 9
    i = i + Val(Mid(CC, 18, 1)) * 7
    i = i + Val(Mid(CC, 19, 1)) * 3
    i = i + Val(Mid(CC, 20, 1)) * 6
    i2 = Int(i / 11)
    i3 = i - (i2 * 11)
    i4 = 11 - i3
    Select Case i4
        Case 11
            i4 = 0
        Case 10
            i4 = 1
    End Select
    DC = DC & i4
    '-- Si llega aqu� ambos figitos de control son correctos
    DigitoControlCorrecto = DC
End Function






