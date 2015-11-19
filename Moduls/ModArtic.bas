Attribute VB_Name = "ModArtic"
' ### [Monica] 29/09/2006
' function que indica si un articulo pertenece a la familia de combustibles

Public Function EsArticuloCombustible(articulo As String) As Boolean
Dim Famia As String
Dim tipoF As String

    EsArticuloCombustible = False
    Famia = ""
    Famia = DevuelveDesdeBD("codfamia", "sartic", "codartic", articulo, "N")
    If Famia = "" Then Exit Function
    tipoF = ""
    tipoF = DevuelveDesdeBD("tipfamia", "sfamia", "codfamia", Famia, "N")
    If tipoF = "" Then Exit Function
    If tipoF = "1" Then EsArticuloCombustible = True

End Function

Public Function EsArticuloDescuento(articulo As String) As Boolean
Dim Famia As String
Dim tipoF As String

    EsArticuloDescuento = False
    Famia = ""
    Famia = DevuelveDesdeBD("codfamia", "sartic", "codartic", articulo, "N")
    If Famia = "" Then Exit Function
    tipoF = ""
    tipoF = DevuelveDesdeBD("tipfamia", "sfamia", "codfamia", Famia, "N")
    If tipoF = "" Then Exit Function
    If tipoF = "2" Then EsArticuloDescuento = True

End Function



Public Function ImpuestoArticulo(articulo As String) As Currency
Dim Sql As String

    ImpuestoArticulo = 0
    Sql = DevuelveDesdeBD("impuesto", "sartic", "codartic", articulo, "N")
    If Sql <> "" Then ImpuestoArticulo = DBLet(CCur(Sql), "N")

End Function

Public Function InsertarFamiliaSiNoExiste(Fam As String) As Boolean
Dim Sql As String

On Error GoTo eInsertarFamiliaSiNoExiste

    InsertarFamiliaSiNoExiste = True
    Sql = ""
    Sql = DevuelveDesdeBD("codfamia", "sfamia", "codfamia", Fam, "N")
    If Sql = "" Then
        Sql = "insert into sfamia (codfamia, nomfamia, tipfamia) values ("
        Sql = Sql & DBSet(Fam, "N") & ",'AUTOMATICA',0)"
        
        Conn.Execute Sql
    End If
    
eInsertarFamiliaSiNoExiste:
    If Err.Number <> 0 Then
        InsertarFamiliaSiNoExiste = False
    End If
End Function

Public Function InsertarArticuloSiNoExiste(Art As String, Fam As String, Nombre As String, precio As String, CodIVA As String) As Boolean
Dim Sql As String
Dim Codmacta As String
Dim CodmacCl As String
Dim vParamAplic As CParamAplic

On Error GoTo eInsertarArticuloSiNoExiste


    Set vParamAplic = New CParamAplic
    
    If vParamAplic.Leer = 0 Then
        Codmacta = vParamAplic.CtaFamDefecto
        CodmacCl = vParamAplic.CtaFamDefecto
        
        InsertarArticuloSiNoExiste = True
        Sql = ""
        Sql = DevuelveDesdeBD("codartic", "sartic", "codartic", Art, "N")
        If Sql = "" Then
                 
            If Not EsArticuloDescuento(Art) And Not EsArticuloCombustible(Art) Then
                Codmacta = vParamAplic.RaizCtaSoc & Format(CInt(Fam), "00000")
                CodmacCl = vParamAplic.RaizCtaCli & Format(CInt(Fam), "00000")
            End If
            
            
            Sql = "insert into sartic (codartic, nomartic, codfamia, codmacta, " & _
                   "codmaccl, codigiva, preventa) values (" & _
                   DBSet(Art, "N") & "," & DBSet(Nombre, "T") & "," & DBSet(Fam, "N") & "," & _
                   DBSet(Codmacta, "T") & "," & DBSet(CodmacCl, "T") & "," & DBSet(CodIVA, "N") & "," & _
                   DBSet(precio, "N") & ")"
    
            Conn.Execute Sql
        End If
    End If
eInsertarArticuloSiNoExiste:
    If Err.Number <> 0 Then
        InsertarArticuloSiNoExiste = False
    End If
End Function

Public Function NombreFichero(path As String) As String
Dim cad As String
Dim cad1 As String
Dim b As Boolean
Dim longitud As Integer
Dim i As Integer
Dim J As Integer

    cad = path
    i = 1
    J = Len(cad)
    b = False
    While Not b
        If InStr(cad, "\") = 0 Then
            b = True
        Else
            cad = Mid(cad, InStr(cad, "\") + 1, J)
            J = Len(cad)
        End If
    Wend
    NombreFichero = cad
End Function

'[Monica]15/02/2011: calculo del importe sigaus de un articulo
Public Function ImpuestoSigausArticulo(articulo As String, cantidad As String) As Currency
Dim Sql As String
Dim Rs As ADODB.Recordset
Dim vCantidad As Currency

    On Error GoTo eImpuestoSigausArticulo
    
    ImpuestoSigausArticulo = 0
    
    Sql = "select pesoart, precsigaus from sartic where codartic = " & DBSet(articulo, "N")
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not Rs.EOF Then
        vCantidad = 0
        If cantidad <> "" Then vCantidad = CCur(cantidad)
        
        ImpuestoSigausArticulo = Round2(vCantidad * DBLet(Rs!pesoart, "N") * DBLet(Rs!precsigaus, "N"), 2)
    End If
    Set Rs = Nothing
    
    Exit Function
    
eImpuestoSigausArticulo:
    MuestraError Err.Number, "Cálculo Impuesto Sigaus", Err.decription
End Function



