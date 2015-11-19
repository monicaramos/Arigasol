Attribute VB_Name = "LibImpresionTicket"
Option Explicit


''David
''Llamara a esta funcion. Si el tipo de documento 32 (tickets) pone impresion directa, lo dejamos como esta, si no...
'' hay que hacer a traves del rpt
'Public Sub ImprimirTicketDirecto(NumTicket As String, FechaTicket As Date, Optional Entregado As Currency, Optional Cambio As Currency)   ' (RAFA/ALZIRA 05092006)
'Dim Directo As Boolean
'Dim cadParam As String
'Dim numParam As Byte
'Dim cadNomRPT As String
'Dim NomImpre As String
'
'    Directo = True
'    If Not PonerParamRPT(32, cadParam, numParam, cadNomRPT, Directo, pPdfRpt) Then Directo = True
'
'
'    ' ----  [07/10/2009] [LAURA] : se poner general para impresion directa y crystal reports
'    ' -- Establecemos la impresora de ticket
'    If vParamTPV.NomImpresora <> "" Then
'        If Printer.DeviceName <> vParamTPV.NomImpresora Then
'            'guardamos la impresora que habia
'            NomImpre = Printer.DeviceName
'            'establecemos la de ticket
'            EstablecerImpresora vParamTPV.NomImpresora
'        End If
'    End If
'    ' ---- []
'
'    If Directo Then
'        '-- Impresion directa
'        ImprimirElTicketDirecto2 NumTicket, FechaTicket, Not vParamTPV.Redondea2, Entregado, Cambio
'    Else
'        'Establecemos la impresora de ticket
''        If vParamTPV.NomImpresora <> "" Then
''            If Printer.DeviceName <> vParamTPV.NomImpresora Then
''                'guardamos la impresora que habia
''                NomImpre = Printer.DeviceName
''                'establecemos la de ticket
''                EstablecerImpresora vParamTPV.NomImpresora
''            End If
''        End If
'
'        '-- Con crystal
'        With frmImprimir
'            .FormulaSeleccion = " {scafac.codtipom} = 'FTI'" & _
'                " and {scafac.numfactu} = " & CStr(NumTicket) & _
'                " and {scafac.fecfactu} = Date(" & Year(FechaTicket) & "," & Month(FechaTicket) & "," & Day(FechaTicket) & ")"
'
'            .OtrosParametros = ""
'            .NumeroParametros = 0
'            .SoloImprimir = True
'            .EnvioEMail = False
'            .Opcion = 93
'            .Titulo = "Ticket"
'            .NombreRPT = cadNomRPT
'            .NombrePDF = pPdfRpt
'            .ConSubInforme = True
'            .Show vbModal
'         End With
'
'
'
'
'        'sI ABRE EL CAJON
'        If vParamTPV.AbreCajon Then ImprimePorLaCom ""
'
'
''        'Volver la impresora a la predeterminada
''        If NomImpre <> "" Then EstablecerImpresora NomImpre
'    End If
'
'
'    ' ----  [07/10/2009] [LAURA] : se poner general para impresion directa y crystal reports
'    ' -- Volver la impresora a la predeterminada
'    If NomImpre <> "" Then EstablecerImpresora NomImpre
'    ' ----- []
'End Sub
'



'Obligo la fecha. Antes NO y la cogia de rsventa
'Public Sub ImprimirTicketDirecto(NumTicket As String, NumAlbTicket1 As String, FechaTicket As Date)  ' (RAFA/ALZIRA 05092006)
Public Sub ImprimirElTicketDirecto2(NumTicket As String, FechaTicket As Date, Precio4Decimales As Boolean, Optional Entregado As Currency, Optional Cambio As Currency)   ' (RAFA/ALZIRA 05092006)
'    Dim NomImpre As String
  '  Dim FechaT As Date
    Dim Rs1 As ADODB.Recordset
    Dim Rs2 As ADODB.Recordset
    Dim rs3 As ADODB.Recordset
    Dim rs4 As ADODB.Recordset
    Dim Sql As String
    Dim lin As String ' línea de impresión
    Dim i As Integer
    Dim N As Integer
    Dim ImporteIva As Currency
    Dim EnEfectivo As Boolean
    
    Dim NomArtic As String
    
On Error GoTo EImpTickD
    'Antes DAVID
'    If FechaTicket = "" Then
'        FechaT = RSVenta!fecventa
'    Else
'        FechaT = CDate(FechaTicket)
'    End If
'       --> Como desde aqui no se ve el rsventa entonces OBLIGAMOS a que se traiga la fecha

'    Stop
   Printer.Font = "Courier New"
    
    
    ' ----  [07/10/2009] [LAURA] : se poner general para impresion directa y crystal reports
'    'Establecemos la impresora de ticket
'    If vParamTPV.NomImpresora <> "" Then
'        If Printer.DeviceName <> vParamTPV.NomImpresora Then
'            'guardamos la impresora que habia
'            NomImpre = Printer.DeviceName
'            'establecemos la de ticket
'            EstablecerImpresora vParamTPV.NomImpresora
'        End If
'    End If
    
    '-- Obtenemos cabeceras y pies en un recordset (rs1)
    Sql = "select scaalb.*, ssocio.nomsocio from scaalb, ssocio where scaalb.codsocio = ssocio.codsocio "
    Sql = Sql & " and scaalb.numalbar = " & DBSet(NumTicket, "N")
    
    Set Rs1 = New ADODB.Recordset
    Rs1.Open Sql, Conn, adOpenForwardOnly
    If Not Rs1.EOF Then
            '-- Consultamos la forma de pago pa 2 cosas
            '   Para imprimirla en el pie y para en el caso de contado mostrar entregado
            '   y cambio.
            Sql = "select * from sforpa where codforpa = " & CStr(Rs1!Codforpa)
            Set rs4 = New ADODB.Recordset
            rs4.Open Sql, Conn, adOpenForwardOnly
            If Not rs4.EOF Then
                '[Monica]04/01/2013: efectivos
                If rs4!TipForpa = 0 Or rs4!TipForpa = 6 Then EnEfectivo = True
            End If
            '-- Impresión de la cabecera
'                Lin = "         1         2         3         4"
'                Printer.Print Lin
'                Lin = "1234567890123456789012345678901234567890"
'                Printer.Print Lin
            
            ' nombre empresa
            lin = LineaCentrada(vEmpresa.nomEmpre)
            If lin <> "" Then Printer.Print lin
            ' domicilio
            lin = LineaCentrada(vEmpresa.DomicilioEmpresa)
            If lin <> "" Then Printer.Print lin
            ' poblacion
            lin = LineaCentrada(vEmpresa.CPostal & " - " & vEmpresa.Poblacion)
            If lin <> "" Then Printer.Print lin
            ' provincia
            lin = LineaCentrada(vEmpresa.Provincia)
            If lin <> "" Then Printer.Print lin
            ' cif empresa
            lin = LineaCentrada("CIF:" & vEmpresa.CifEmpresa)
            If lin <> "" Then Printer.Print lin
            Printer.Print ""
            
            
            lin = CuadraParteI(20, "Ticket: " & Format(Rs1!numalbar, "0000000")) & _
                  CuadraParteD(20, "Fecha: " & Format(Rs1!fecAlbar, "dd/mm/yyyy"))
            Printer.Print lin
            
            lin = CuadraParteI(20, "E.S.: " & vParamAplic.Cim) & _
                  CuadraParteD(20, "Hora:       " & Format(Rs1!horalbar, "hh:mm"))
            Printer.Print lin
            
            lin = CuadraParteI(40, "Cliente: " & Format(Rs1!codsocio, "000000") & _
                  "  " & Rs1!NomSocio)
            Printer.Print lin
            
            lin = CuadraParteI(40, "Matrícula: " & Rs1!matricul)
            Printer.Print lin
            
            lin = ""
            Printer.Print lin
            
            lin = CuadraParteI(40, "Forma Pago:" & Format(Rs1!Codforpa, "000") & " " & rs4!nomforpa)
            Printer.Print lin
            
            lin = String(40, "-")
            Printer.Print lin
            lin = CuadraParteI(16, "Producto") & _
                    CuadraParteD(6, " Cant") & _
                    CuadraParteD(8, "PVP") & _
                    CuadraParteD(10, "Importe")
            Printer.Print lin
            lin = String(40, "-")
            Printer.Print lin
            
            NomArtic = DevuelveValor("select nomartic from sartic where codartic = " & DBSet(Rs1!codartic, "N"))
            
            lin = CuadraParteI(16, Mid(NomArtic, 1, 16)) & _
                    CuadraParteD(6, Format(Rs1!cantidad, "##0.00")) & _
                    CuadraParteD(8, Format(Rs1!preciove, "##0.0000")) & _
                    CuadraParteD(10, Format(Rs1!importel, "###,##0.00"))
            Printer.Print lin
            
            lin = String(40, "-")
            Printer.Print lin
            
            '-- Impresion del total
            Printer.Print String(40, " ")
            lin = CuadraParteI(20, "Total Ticket: ") & CuadraParteD(20, Format(Rs1!importel, "###,###,#0.00"))
            Printer.Print lin
            
            Printer.Print ""
            lin = LineaCentrada("IVA INCLUIDO")
            Printer.Print lin
            
            For i = 1 To 8
                Printer.Print String(40, " ")
            Next i
            

'            Printer.Print Chr$(29) & Chr$(86) & "0"
'            Printer.Print Chr$(27) & "i"
            
            
'            Dim Puerto As String
'            Dim nFicSalCajon As Integer
'
'            Puerto = "LPT1"
'            nFicSalCajon = FreeFile
'
'            Open Puerto For Output As #nFicSalCajon
'                Print #nFicSalCajon, Chr$(29); Chr$(86); "0"
'
'            Close nFicSalCajon

            
'            Printer.Print Chr$(29); Chr$(86); Chr$(0) 'arigasol2
            
            
'            Printer.Print Chr$(29) & Chr$(86) & Chr$(0)  'arigasol3
            
'            Dim Puerto As String            ' arigasol6
'            Dim nFicSalCajon As Integer
'            Puerto = "LPT1"
'            nFicSalCajon = FreeFile
'            Open Puerto For Output As #nFicSalCajon
'                Print #nFicSalCajon, Chr$(29); Chr$(86); Chr$(0)
'            Close nFicSalCajon
            
            '-- Fin de impresión
            Printer.NewPage
            Printer.EndDoc


'            Printer.Print Chr$(29); Chr$(86); Chr$(0) 'arigasol1

'            Printer.Print Chr$(29) & Chr$(86) & Chr$(0)  'arigasol4

'            Dim Puerto As String            ' arigasol5
'            Dim nFicSalCajon As Integer
'            Puerto = "LPT1"
'            nFicSalCajon = FreeFile
'            Open Puerto For Output As #nFicSalCajon
'                Print #nFicSalCajon, Chr$(29); Chr$(86); Chr$(0)
'            Close nFicSalCajon

            Dim Puerto As String            ' arigasol7  **********CORRECTA
            Dim nFicSalCajon As Integer
            Puerto = "LPT1"
            nFicSalCajon = FreeFile
            Open Puerto For Output As #nFicSalCajon
                Print #nFicSalCajon, Chr$(27); "i"
            Close nFicSalCajon
'**********
    Else
        MsgBox "No se ha encontrado el ticket " & CStr(NumTicket) & " de " & Format(FechaTicket, "dd/mm/yyyy"), vbCritical
    End If
    
    Rs1.Close
    
    ' ----  [07/10/2009] [LAURA] : se poner general para impresion directa y crystal reports
'    'Volver la impresora a la predeterminada
'    EstablecerImpresora NomImpre
    ' ----  []
    
    Exit Sub
EImpTickD:
    MsgBox Err.Description & " (" & Err.Number & ")", vbCritical, "Imprimir ticket."
End Sub


Private Sub ImprimePorLaCom(cadena As String)
    On Error GoTo EI
    
    Dim nFicSalCajon As Integer
    Dim Puerto As String
    
    'Marzo 2011
    'Puerto = "COM1"
'    Puerto = "COM" & vParamTPV.ComImpresora
    nFicSalCajon = FreeFile
    
    Open Puerto For Output As #nFicSalCajon
    'If Check1.Value = 1 Then
        Print #nFicSalCajon, Chr$(27); "p"; Chr$(0); Chr$(25); Chr$(250)
    'Else
    '    Print #nFicSalCajon, Cadena
    'End If
    
    '- corta papel
        '        Print #IMPRESORA, Chr$(29) + Chr$(86) + "0"
    
    Close nFicSalCajon
    
    Exit Sub
EI:
    cadena = "Error en COM: " & vbCrLf & vbCrLf & Err.Description
    MsgBox cadena, vbCritical
End Sub


Private Sub CortaPapel()
    Printer.Print Chr(29) & Chr(56) & Chr(49)
'    Printer.EndDoc
End Sub




Private Function LineaCentrada(lin As String) As String
    Dim queda As Integer
    Dim parte As Integer
    queda = 40 - Len(lin)
    parte = queda / 2
    If parte Then
        LineaCentrada = String(parte, " ") & lin & String(queda - parte, " ")
    Else
        LineaCentrada = lin
    End If
End Function

Private Function CuadraParteD(longitud As Integer, cadena As String) As String
    CuadraParteD = Right(String(longitud, " ") & cadena, longitud)
End Function

Private Function CuadraParteI(longitud As Integer, cadena As String) As String
    CuadraParteI = Left(cadena & String(longitud, " "), longitud)
End Function

