Attribute VB_Name = "modExpgrupo"
Option Explicit

' +-+- torna TRUE si queden places lliures per al expedient que se li passa +-+-
Public Function QuedenPlaces(numexped As String, codEmpre As String, ind As Integer, dob As Integer, tri As Integer, qua As Integer, ByRef numerpax As Integer, ByRef ocupades As Integer) As Boolean
'Dim demanades As Integer, numerpax As Integer
Dim demanades As Integer
'Dim ocupades As Integer, lliures As Integer
Dim lliures As Integer
Dim nmaxhab1 As Integer, nmaxhab2 As Integer, nmaxhab3 As Integer, nmaxhab4 As Integer
Dim ocup1 As Integer, ocup2 As Integer, ocup3 As Integer, ocup4 As Integer
Dim RS As ADODB.Recordset
Dim SQL As String
    
    demanades = ind + (dob * 2) + (tri * 3) + (qua * 4)

    SQL = "SELECT * FROM expgrupo WHERE numexped = " & numexped & " and codempre=" & codEmpre

    Set RS = New ADODB.Recordset
    RS.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText

    numerpax = RS.Fields!numerpax
    nmaxhab1 = RS.Fields!nmaxhab1
    nmaxhab2 = RS.Fields!nmaxhab2
    nmaxhab3 = DBLet(RS.Fields!nmaxhab3, "N")
    nmaxhab4 = DBLet(RS.Fields!nmaxhab4, "N")
    
    RS.Close
    Set RS = Nothing
    
    ocup1 = DevuelveDesdeBDnew2(1, "COUNT(*)", "viagrupc", "numexped|codempre|tiphabit|", numexped & "|" & codEmpre & "|1|", "N|N|N|", 3)
    ocup2 = DevuelveDesdeBDnew2(1, "COUNT(*)", "viagrupc", "numexped|codempre|tiphabit|", numexped & "|" & codEmpre & "|2|", "N|N|N|", 3)
    ocup3 = DevuelveDesdeBDnew2(1, "COUNT(*)", "viagrupc", "numexped|codempre|tiphabit|", numexped & "|" & codEmpre & "|3|", "N|N|N|", 3)
    ocup4 = DevuelveDesdeBDnew2(1, "COUNT(*)", "viagrupc", "numexped|codempre|tiphabit|", numexped & "|" & codEmpre & "|4|", "N|N|N|", 3)
    
    ocupades = ocup1 + ocup2 + ocup3 + ocup4
    lliures = numerpax - ocupades
    
    If (demanades > lliures) Then
        MsgBox "Ha seleccionado más plazas de las disponibles, sólo quedan " & lliures, vbInformation
        QuedenPlaces = False
        Exit Function
    End If

    If ind > (nmaxhab1 - ocup1) Then
        MsgBox "Ha seleccionado más habitaciones individuales de las disponibles, sólo quedan " & (nmaxhab1 - ocup1), vbInformation
        QuedenPlaces = False
        Exit Function
    ElseIf (dob) > (nmaxhab2 - (ocup2 / 2)) Then
        MsgBox "Ha seleccionado más habitaciones dobles de las disponibles, sólo quedan " & (nmaxhab2 - (ocup2 / 2)), vbInformation
        QuedenPlaces = False
        Exit Function
    ElseIf (tri) > (nmaxhab3 - (ocup3 / 3)) Then
        MsgBox "Ha seleccionado más habitaciones triples de las disponibles, sólo quedan " & (nmaxhab3 - (ocup3 / 3)), vbInformation
        QuedenPlaces = False
        Exit Function
    ElseIf (qua) > (nmaxhab4 - (ocup4 / 4)) Then
        MsgBox "Ha seleccionado más habitaciones cuádruples de las disponibles, sólo quedan " & (nmaxhab4 - (ocup4 / 4)), vbInformation
        QuedenPlaces = False
        Exit Function
    End If

    QuedenPlaces = True

End Function

Public Sub OrdSeients(ByRef a_numasien() As Variant, numexped As String, codEmpre As String, numerpax As Integer, ocupades As Integer)
Dim a_total() As Variant, a_ocupats() As Variant, a_lliures() As Variant
Dim i As Integer, j As Integer
Dim RS As ADODB.Recordset
Dim SQL As String

    ReDim Preserve a_total(numerpax - 1)
    If ocupades > 0 Then
        ReDim Preserve a_ocupats(ocupades - 1)
    End If
    ReDim Preserve a_lliures(numerpax - ocupades - 1)
    ReDim Preserve a_numasien(numerpax - ocupades - 1)

    For i = 0 To (numerpax - 1)
        a_total(i) = 1 'array en tants elements com seients hi han, tindra 1/0 per a indicar si esta lliure o no, en un principi tots lliures
    Next i

    'consulte els seients ocupats
    SQL = "SELECT * FROM viagrupc WHERE numexped = " & numexped & " AND codempre= " & codEmpre & " ORDER BY numasien"

    Set RS = New ADODB.Recordset
    RS.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    
    If Not RS.EOF Then
        For i = 0 To ocupades - 1
            a_ocupats(i) = RS.Fields!numasien
            RS.MoveNext
        Next i
    End If
    
    RS.Close
    Set RS = Nothing
    
    For i = 0 To ocupades - 1
        a_total(a_ocupats(i) - 1) = 0
    Next i

    j = 0 'index de a_lliures
    For i = 0 To numerpax - 1
        If a_total(i) = 1 Then
            a_lliures(j) = i + 1
            j = j + 1
        End If
    Next i
    
    'ara ordenaré el seients, mentri mentres, el torne conforme està
    a_numasien = a_lliures
    
End Sub


Public Sub InsSeients(a_numasien() As Variant, numexped As String, codEmpre As String, ind As Integer, dob As Integer, tri As Integer, qua As Integer, codClien As String, forpa As String, reser As String, localiza As String)
Dim demanades As Integer, relacion As String
Dim i As Integer, j As Integer
Dim a_idhabita() As Variant
Dim a_tiphabit() As Variant
Dim c1 As Integer, c2 As Integer, c3 As Integer, c4 As Integer
Dim SQL As String

    demanades = ind + (dob * 2) + (tri * 3) + (qua * 4)
    
    'redimensione esl vectors
    ReDim Preserve a_idhabita(demanades - 1)
    ReDim Preserve a_tiphabit(demanades - 1)
    
    'averigüe quin serà el valor de relacion
    relacion = SugerirCodigoSiguienteStr("viagrupc", "relacion", "numexped=" & numexped & " and codempre=" & codEmpre)

    j = 0 'index generals per als vectors

    'prepare els vectors
    For i = 1 To ind
        a_idhabita(j) = 1
        a_tiphabit(j) = 1
        j = j + 1
    Next i
    
    For i = 1 To dob
        a_idhabita(j) = 1
        a_idhabita(j + 1) = 2
        a_tiphabit(j) = 2
        a_tiphabit(j + 1) = 2
        j = j + 2
    Next i
    
    For i = 1 To tri
        a_idhabita(j) = 1
        a_idhabita(j + 1) = 2
        a_idhabita(j + 2) = 3
        a_tiphabit(j) = 3
        a_tiphabit(j + 1) = 3
        a_tiphabit(j + 2) = 3
        j = j + 3
    Next i
    
    For i = 1 To qua
        a_idhabita(j) = 1
        a_idhabita(j + 1) = 2
        a_idhabita(j + 2) = 3
        a_idhabita(j + 3) = 4
        a_tiphabit(j) = 4
        a_tiphabit(j + 1) = 4
        a_tiphabit(j + 2) = 4
        a_tiphabit(j + 3) = 4
        j = j + 4
    Next i
    
    ' bucle per a insertar els viatgers
    For i = 0 To (demanades - 1)
        Select Case a_tiphabit(i)
            Case 1 'individual
                c1 = 0
                c2 = 0
                c3 = 0
                c4 = 0
            Case 2 'doble
                Select Case a_idhabita(i)
                    Case 1
                        c1 = a_numasien(i + 1)
                        c2 = 0
                        c3 = 0
                        c4 = 0
                    Case 2
                        c1 = a_numasien(i - 1)
                        c2 = 0
                        c3 = 0
                        c4 = 0
                End Select
            Case 3 'triple
                Select Case a_idhabita(i)
                    Case 1
                        c1 = a_numasien(i + 1)
                        c2 = a_numasien(i + 2)
                        c3 = 0
                        c4 = 0
                    Case 2
                        c1 = a_numasien(i - 1)
                        c2 = a_numasien(i + 1)
                        c3 = 0
                        c4 = 0
                    Case 3
                        c1 = a_numasien(i - 2)
                        c2 = a_numasien(i - 1)
                        c3 = 0
                        c4 = 0
                End Select
            Case 4 'quádruple
                Select Case a_idhabita(i)
                    Case 1
                        c1 = a_numasien(i + 1)
                        c2 = a_numasien(i + 2)
                        c3 = a_numasien(i + 3)
                        c4 = 0
                    Case 2
                        c1 = a_numasien(i - 1)
                        c2 = a_numasien(i + 1)
                        c3 = a_numasien(i + 2)
                        c4 = 0
                    Case 3
                        c1 = a_numasien(i - 2)
                        c2 = a_numasien(i - 1)
                        c3 = a_numasien(i + 1)
                        c4 = 0
                    Case 4
                        c1 = a_numasien(i - 3)
                        c2 = a_numasien(i - 2)
                        c3 = a_numasien(i - 1)
                        c4 = 0
                End Select
        End Select 'end del case de a_tiphabit
    
        'INSERT INTO viagrupc VALUES("10500002", "1", "2005-02-02", "1", "1", "1", "", "", "0", "nom", "ape", "", "", NULL, "", "", NULL, NULL, NULL, "1", "1", "1", "0", "0", "0", "0", "0", "0", "0", "0000-00-00", "1", "0.00", NULL, "2", "0.00");
        '(numexped,codempre,numasien,fechalta,codemple,codagenc,codclien,reservad,localiza,situacio,nompasaj,apepasaj,nifpasaj,dirpasaj,codpobla,codposta,telefono,edadviaj,ciurecog,lugrecog,relacion,tiphabit,idhabita,compar_1,compar_2,compar_3,compar_4,bonoimpr,tipofact,numfactu,fecfactu,codforpa,comisclt,impasien,observac)
        SQL = "INSERT INTO viagrupc (numexped,codempre,numasien,fechalta,codemple,codagenc,codclien,reservad,localiza,situacio,nompasaj,apepasaj,nifpasaj,dirpasaj,codpobla,codposta,telefono,edadviaj,ciurecog,lugrecog,relacion,tiphabit,idhabita,compar_1,compar_2,compar_3,compar_4,bonoimpr,tipofact,numfactu,fecfactu,codforpa,comisclt,impasien,observac)"
        SQL = SQL & " VALUES(" & numexped & ", " & codEmpre & "," & a_numasien(i) & ", " & DBSet(Now(), "F", "N") & ", " & vSesion.Empleado & ", " & vSesion.Agencia & "," & DBSet(codClien, "N") & "," & DBSet(reser, "T") & ", " & DBSet(localiza, "T") & ", 0, ""NOMBRE_TEMPORAL_" & (i + 1) & """, ""APELLIDOS_TEMPORAL_" & (i + 1) & """, """", """", NULL, """", """", NULL, NULL, NULL, " & relacion & ", " & a_tiphabit(i) & ", " & a_idhabita(i) & ", " & c1 & ", " & c2 & ", " & c3 & ", " & c4 & ", ""0"", ""0"", ""0"", ""0000-00-00"", " & forpa & ", ""0.00"",""0.00"", NULL " & ")"
        Conn.Execute SQL
    Next i
    
    
End Sub
