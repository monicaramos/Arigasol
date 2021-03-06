Attribute VB_Name = "ModFacturacion"
' Modulo en donde se encuentran los procedimintos para la facturacion
'
'Dim db As BaseDatos
Dim Rs As ADODB.Recordset
Dim ImpFactu As Currency
Dim TotalImp As Currency

Dim TotalImpSigaus As Currency '[Monica]15/02/2011 Impuesto de lubricantes

Dim numser As String
Dim dc As Dictionary
Dim baseimpo As Dictionary


Public Function TraspasoHistoricoFacturas(db As BaseDatos, SQL As String, desde As String, hasta As String, ByRef Pb1 As ProgressBar) As Boolean
    
    Dim importel As Currency
    Dim impbase As Currency
    Dim actFactura As Long
    Dim antfactura As Long
    Dim AntFecha As Date
    Dim AntSocio As Long
    Dim AntForpa As Integer
    Dim HayReg As Boolean
    
    Dim Sql1 As String
    
    Dim NumError As Long

    On Error GoTo eTraspasoHistoricoFacturas
    
'    Set db = New BaseDatos
'
'    db.abrir "arigasol", "root", "aritel"
'    db.Tipo = "MYSQL"
        
    Set baseimpo = New Dictionary
      
    ' abrimos la transaccion
    db.AbrirTrans
      
      ' traemos el numero de serie de la factura de tipo FAC de la tabla stipom
      numser = ""
      numser = DevuelveDesdeBD("letraser", "stipom", "codtipom", "FAT", "T")
      
      TotalImp = 0
      TotalImpSigaus = 0
      Set Rs = db.cursor(SQL)
      HayReg = False
      NumError = 0
      If Not Rs.EOF Then
          Rs.MoveFirst
          antfactura = Rs!numfactu
          AntFecha = Rs!fecAlbar
          AntSocio = Rs!codsocio
          AntForpa = Rs!Codforpa
          
          While Not Rs.EOF And NumError = 0
              actFactura = Rs!numfactu
              HayReg = True
              If actFactura <> antfactura Then ' after group of numfactu
                 If NumError = 0 Then NumError = InsertCabe(db, baseimpo, antfactura, AntFecha, AntSocio, AntForpa, 0)
                 Set baseimpo = Nothing
                 Set baseimpo = New Dictionary
                 TotalImp = 0
                 antfactura = actFactura
                 AntFecha = Rs!fecAlbar
                 AntSocio = Rs!codsocio
                 AntForpa = Rs!Codforpa
              End If
              '-------
              ' tenemos que calcular el impuesto multiplicando cantidad de linea por impuesto por articulo
              importel = DBLet(Rs!Impuesto, "N") ' Comprueba si es nulo y lo pone a 0 o ""
              
              If EsArticuloCombustible(Rs!codartic) Then
                TotalImp = TotalImp + Round2((Rs!cantidad * importel), 2)
              End If
              baseimpo(Val(Rs!CodigIVA)) = DBLet(baseimpo(Val(Rs!CodigIVA)), "N") + DBLet(Rs!importel, "N")
              
              TotalImpSigaus = TotalImpSigaus + ImpuestoSigausArticulo(CStr(DBLet(Rs!codartic, "N")), CStr(DBLet(Rs!cantidad, "N")))
              
              If NumError = 0 Then NumError = InsertLinea(db, Rs)
              
              If NumError = 0 Then
                    Pb1.Value = Pb1.Value + 1
                    Pb1.Refresh
                    
                    Rs.MoveNext
              End If
          Wend
          If HayReg And NumError = 0 Then NumError = InsertCabe(db, baseimpo, actFactura, AntFecha, AntSocio, AntForpa, 0)


          ' hacemos el borrado masivo de albaranes de las los albaranes
          If NumError = 0 Then NumError = BorradoAlbaranes(db, desde, hasta)

          ' aprovechamos para borrar todas las pruebas de manguera
          If NumError = 0 Then NumError = BorradoAlbaranesPrueba(db, desde, hasta)

        End If
        
    Set Rs = Nothing
    If NumError <> 0 Then Err.Raise NumError
        
eTraspasoHistoricoFacturas:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Error en el traspaso al hist�rico. Llame a soporte." & vbCrLf & vbCrLf & MensError
        db.RollbackTrans
        TraspasoHistoricoFacturas = False
        Pb1.visible = False
    Else
        db.CommitTrans
        TraspasoHistoricoFacturas = True
    End If
End Function

'Insertar Cabecera de factura
Public Function InsertCabe(ByRef db As BaseDatos, ByRef dc As Dictionary, numfactu As Long, Fecha As Date, socio As Long, Forpa As Integer, Tipo As Byte, Optional Contabilizada As Boolean, Optional SinIva As Boolean, Optional Dpto As Integer) As Long     ', db As Database)
' tipo 0 en la schfac
' tipo 1 en la schfacr

    Dim I As Integer
    Dim Imptot(2)
    Dim Tipiva(2)
    Dim Impbas(2)
    Dim ImpIva(2)
    Dim PorIva(2)
    Dim TotFac
    Dim SQL As String
    Dim NumCoop As String
    
    '05012007
    On Error GoTo eInsertCabe
    MensError = ""
    ' inicializamos los importes de los totales de la cabecera
    TotFac = 0
    For I = 0 To 2
         Tipiva(I) = Null
         Imptot(I) = Null
         Impbas(I) = Null
         ImpIva(I) = Null
         PorIva(I) = Null
    Next I
    
    For I = 0 To dc.Count - 1
        If I <= 2 Then '  And i = 0 Then
            If SinIva Then
                If I = 0 Then
                    Tipiva(0) = vParamAplic.TipoIvaExento
                    Imptot(0) = dc.Items(0)
                    PorIva(0) = DevuelveDesdeBDNew(cConta, "tiposiva", "porceiva", "codigiva", CStr(Tipiva(0)), "N")
                    Impbas(0) = Round2(Imptot(0) / (1 + (PorIva(0) / 100)), 2)
                    ImpIva(0) = Imptot(0) - Impbas(0)
                    TotFac = Imptot(0)
                Else
                    Imptot(0) = Imptot(0) + dc.Items(I)
                    PorIva(0) = DevuelveDesdeBDNew(cConta, "tiposiva", "porceiva", "codigiva", CStr(Tipiva(0)), "N")
                    Impbas(0) = Round2(Imptot(0) / (1 + (PorIva(0) / 100)), 2)
                    ImpIva(0) = Imptot(0) - Impbas(0)
                    TotFac = Imptot(0)
                End If
            Else
                '[Monica]04/02/2013: si el importe es 0 no lo insertamos
                '                    solo si no es el primero
                If I = 0 Then
                    Tipiva(I) = dc.Keys(I)
                    Imptot(I) = dc.Items(I)
                    PorIva(I) = DevuelveDesdeBDNew(cConta, "tiposiva", "porceiva", "codigiva", CStr(Tipiva(I)), "N")
                    Impbas(I) = Round2(Imptot(I) / (1 + (PorIva(I) / 100)), 2)
                    ImpIva(I) = Imptot(I) - Impbas(I)
                    TotFac = TotFac + Imptot(I)
                Else
                    If dc.Items(I) = 0 Then
                        I = I + 1
                        If I = 3 Then
                            Tipiva(I - 1) = dc.Keys(I)
                            Imptot(I - 1) = dc.Items(I)
                            PorIva(I - 1) = DevuelveDesdeBDNew(cConta, "tiposiva", "porceiva", "codigiva", CStr(Tipiva(I)), "N")
                            Impbas(I - 1) = Round2(Imptot(I) / (1 + (PorIva(I) / 100)), 2)
                            ImpIva(I - 1) = Imptot(I) - Impbas(I)
                            TotFac = TotFac + Imptot(I)
                            
                            Exit For
                        End If
                    Else
                        Tipiva(I) = dc.Keys(I)
                        Imptot(I) = dc.Items(I)
                        PorIva(I) = DevuelveDesdeBDNew(cConta, "tiposiva", "porceiva", "codigiva", CStr(Tipiva(I)), "N")
                        Impbas(I) = Round2(Imptot(I) / (1 + (PorIva(I) / 100)), 2)
                        ImpIva(I) = Imptot(I) - Impbas(I)
                        TotFac = TotFac + Imptot(I)
                    End If
                End If
            End If
        End If
    Next I
    '    TotFac = TotFac - totalimp
    
    NumCoop = DevuelveDesdeBD("codcoope", "ssocio", "codsocio", CStr(socio), "N")
    
    If Tipo = 0 Then
        SQL = "INSERT into schfac "
    
        SQL = SQL & "(letraser, numfactu, fecfactu, codsocio, codcoope, " & _
               "codforpa, baseimp1, baseimp2, baseimp3, impoiva1, " & _
               "impoiva2, impoiva3, tipoiva1, tipoiva2, tipoiva3, " & _
               "porciva1, porciva2, porciva3, totalfac, impuesto, impuesigaus, " & _
               "intconta, coddepar) " & _
               "values " & _
               "(" & db.Texto(numser) & "," & db.numero(numfactu) & "," & db.Fecha(Fecha) & "," & db.numero(socio) & "," & db.numero(NumCoop) & "," & _
               db.numero(Forpa) & "," & db.numero(Impbas(0)) & "," & db.numero(Impbas(1)) & "," & db.numero(Impbas(2)) & "," & db.numero(ImpIva(0)) & "," & _
               db.numero(ImpIva(1)) & "," & db.numero(ImpIva(2)) & "," & db.numero(Tipiva(0)) & "," & db.numero(Tipiva(1)) & "," & db.numero(Tipiva(2)) & "," & _
               db.numero(PorIva(0)) & "," & db.numero(PorIva(1)) & "," & db.numero(PorIva(2)) & "," & db.numero(TotFac) & "," & db.numero(TotalImp) & "," & db.numero(TotalImpSigaus) & ","
    
    Else
        SQL = "INSERT into schfacr "
    
        SQL = SQL & "(letraser, numfactu, fecfactu, codsocio, codcoope, " & _
               "codforpa, baseimp1, baseimp2, baseimp3, impoiva1, " & _
               "impoiva2, impoiva3, tipoiva1, tipoiva2, tipoiva3, " & _
               "porciva1, porciva2, porciva3, totalfac, impuesto, " & _
               "intconta) " & _
               "values " & _
               "(" & db.Texto(numser) & "," & db.numero(numfactu) & "," & db.Fecha(Fecha) & "," & db.numero(socio) & "," & db.numero(NumCoop) & "," & _
               db.numero(Forpa) & "," & db.numero(Impbas(0)) & "," & db.numero(Impbas(1)) & "," & db.numero(Impbas(2)) & "," & db.numero(ImpIva(0)) & "," & _
               db.numero(ImpIva(1)) & "," & db.numero(ImpIva(2)) & "," & db.numero(Tipiva(0)) & "," & db.numero(Tipiva(1)) & "," & db.numero(Tipiva(2)) & "," & _
               db.numero(PorIva(0)) & "," & db.numero(PorIva(1)) & "," & db.numero(PorIva(2)) & "," & db.numero(TotFac) & "," & db.numero(TotalImp) & ","
        
    End If


    If Contabilizada Then
        SQL = SQL & "1"
    Else
        SQL = SQL & "0"
    End If
    
    '[Monica]03/05/2019: hay que guardarse el departamento
    If Tipo = 0 Then
        SQL = SQL & "," & DBSet(Dpto, "N", "N") & ")"
    Else
        SQL = SQL & ")"
    End If
    
    InsertCabe = db.ejecutar(SQL)

eInsertCabe:
    If Err.Number <> 0 Or InsertCabe <> 0 Then
        MensError = "Error en la inserci�n en schfac de la factura " & numfactu & " del socio " & socio
        If InsertCabe = 0 Then InsertCabe = 1
    End If
End Function

'Insertar Linea de factura
Public Function InsertLinea(db As BaseDatos, Rs As ADODB.Recordset) As Long  ' , db As Database) As Boolean

    Dim SQL As String
    Dim ImpLinea As Currency
    
'05012007
On Error GoTo eInsertLinea
    MensError = ""
    
        SQL = "INSERT into slhfac " & _
           "(letraser, numfactu, fecfactu, numlinea, numalbar, " & _
           "fecalbar, horalbar, codturno, numtarje, codartic, " & _
           "cantidad, preciove, implinea, kilometros, dtoalvic, importevale) " & _
           "values " & _
           "(" & db.Texto(numser) & "," & db.numero(Rs!numfactu) & "," & db.Fecha(Rs!fecAlbar) & "," & db.numero(Rs!NumLinea) & "," & db.Texto(Rs!numalbar) & "," & _
           db.Fecha(Rs!fecAlbar) & "," & db.fechahora(Rs!fecAlbar & " " & Format(Rs!horalbar, "hh:mm:ss")) & "," & db.numero(Rs!codTurno) & "," & db.numero(Rs!Numtarje) & "," & db.numero(Rs!codartic) & "," & _
           db.numero(Rs!cantidad) & "," & db.numero(Rs!preciove) & "," & db.numero(Rs!importel) & "," & _
           db.numero(Rs!Kilometros) & "," & db.numero(Rs!dtoalvic) & "," & db.numero(Rs!ImporteVale) & ")"
    
    InsertLinea = db.ejecutar(SQL)
    
eInsertLinea:
    If Err.Number <> 0 Or InsertLinea <> 0 Then
        MensError = "Se ha producido un error en la inserci�n de la linea de factura correspondiente al albaran " & Rs!numalbar
        If InsertLinea = 0 Then InsertLinea = 1
    End If
    
End Function


Public Function ExisteEnHistorico(cDesde As String, cHasta As String, ctipo As String) As Boolean
Dim SQL As String
Dim Tipo As String

    ExisteEnHistorico = False
    
    SQL = "select count(*) from slhfac, scaalb where letraser = " & DBSet(Tipo, "T") & " and " & _
          " slhfac.numfactu= scaalb.numfactu and slhfac.numlinea = scaalb.numlinea "
    
    If cDesde <> "" Then SQL = SQL & " and scaalb.fecalbar >= '" & Format(cDesde, FormatoFecha) & "' "
    If cHasta <> "" Then SQL = SQL & " and scaalb.fecalbar <= '" & Format(cHasta, FormatoFecha) & "' "

    ExisteEnHistorico = (TotalRegistros(SQL) <> 0)
    
End Function


Public Sub RecalculoBasesIvaFactura(ByRef Rs As ADODB.Recordset, ByRef Imptot As Variant, ByRef Tipiva As Variant, ByRef Impbas As Variant, ByRef ImpIva As Variant, ByRef PorIva As Variant, ByRef TotFac As Currency, ByRef totimp As Currency, ByRef totimpSigaus As Currency)

    Dim I As Integer
    Dim SQL As String
    Dim baseimpo As Dictionary
    Dim CodIVA As Integer

    Set baseimpo = New Dictionary

    ' inicializamos los importes de los totales de la cabecera
    TotFac = 0
    totimp = 0
    totimpSigaus = 0
    For I = 0 To 2
         Tipiva(I) = 0
         Imptot(I) = 0
         Impbas(I) = 0
         ImpIva(I) = 0
         PorIva(I) = 0
    Next I

    ' recorremos todas las lineas de la factura
    If Not Rs.EOF Then Rs.MoveFirst
    While Not Rs.EOF
        If EsArticuloCombustible(CStr(Rs!codartic)) Then
            Impuesto = ImpuestoArticulo(Rs!codartic)
            
            totimp = totimp + Round2(Rs!cantidad * Impuesto, 2)
        End If
        
        totimpSigaus = totimpSigaus + ImpuestoSigausArticulo(CStr(DBLet(Rs!codartic, "N")), CStr(DBLet(Rs!cantidad, "N")))
        '[Monica]25/07/2013: letra de serie
        'If Rs!letraser = vParamAplic.LetraInt Then
        If EsInterna(Rs!letraser) Then
            CodIVA = vParamAplic.TipoIvaExento
        Else
            CodIVA = DevuelveDesdeBD("codigiva", "sartic", "codartic", DBLet(Rs!codartic), "N")
        End If
        
        baseimpo(Val(CodIVA)) = DBLet(baseimpo(Val(CodIVA)), "N") + DBLet(Rs!ImpLinea, "N")

        Rs.MoveNext
    Wend

    For I = 0 To baseimpo.Count - 1
        If I <= 2 Then
            Tipiva(I) = baseimpo.Keys(I)
            Imptot(I) = baseimpo.Items(I)
' antes
'            PorIva(i) = DevuelveDesdeBDNew(cConta, "tiposiva", "porceiva", "codigiva", CStr(Tipiva(i)), "N")
'            impiva(i) = DBLet(Round2(Imptot(i) * PorIva(i) / 100, 2), "N")
'            Impbas(i) = Imptot(i) - impiva(i)
'            TotFac = TotFac + Imptot(i)
' ahora
            PorIva(I) = DevuelveDesdeBDNew(cConta, "tiposiva", "porceiva", "codigiva", CStr(Tipiva(I)), "N")
            Impbas(I) = Round2(Imptot(I) / (1 + (PorIva(I) / 100)), 2)
            ImpIva(I) = Imptot(I) - Impbas(I)
            TotFac = TotFac + Imptot(I)
        
        
        End If
    Next I

End Sub

Public Function InsertaLineaFactura(ByRef db As BaseDatos, Rs As ADODB.Recordset, numser As String, numFac As Long, fecFac As Date, Linea As Integer, Tipo As Byte) As Long
' tipo = 0 facturacion
' tipo = 1 facturacion ajena

    Dim SQL As String
    Dim ImpLinea As Currency
    
    On Error GoTo eInsertaLineaFactura
    MensError = ""
    
    If Tipo = 0 Then
        SQL = "INSERT into slhfac "
    
       '[Monica]24/06/2013: introducimos los kilometros
            '[Monica]24/08/2015: introducimos el descuento de alvic para el regaixo
        SQL = SQL & "(letraser, numfactu, fecfactu, numlinea, numalbar, " & _
              "fecalbar, horalbar, codturno, numtarje, codartic, " & _
              "cantidad, preciove, implinea, matricul, precioinicial, kilometros, dtoalvic, importevale ) " & _
              "values " & _
              "(" & db.Texto(numser) & "," & db.numero(numFac) & "," & db.Fecha(fecFac) & "," & db.numero(Linea) & "," & db.Texto(Rs!numalbar) & "," & _
              db.Fecha(Rs!fecAlbar) & "," & db.fechahora(Rs!fecAlbar & " " & Format(Rs!horalbar, "hh:mm:ss")) & "," & db.numero(Rs!codTurno) & "," & db.numero(Rs!Numtarje) & "," & db.numero(Rs!codartic) & "," & _
              db.numero(Rs!cantidad) & "," & db.numero(Rs!preciove) & "," & db.numero(Rs!importel) & "," & db.Texto(Rs!matricul) & "," & db.numero(Rs!precioinicial) & "," & _
              db.numero(Rs!Kilometros) & "," & db.numero(Rs!dtoalvic) & "," & db.numero(Rs!ImporteVale) & ")"
    Else
        SQL = "INSERT into slhfacr "
    
    
        '[Monica]24/06/2013: introducimos los kilometros
            '[Monica]24/08/2015: introducimos el descuento de alvic para el regaixo
        SQL = SQL & "(letraser, numfactu, fecfactu, numlinea, numalbar, " & _
              "fecalbar, horalbar, codturno, numtarje, codartic, " & _
              "cantidad, preciove, implinea, matricul, kilometros, dtoalvic, importevale) " & _
              "values " & _
              "(" & db.Texto(numser) & "," & db.numero(numFac) & "," & db.Fecha(fecFac) & "," & db.numero(Linea) & "," & db.Texto(Rs!numalbar) & "," & _
              db.Fecha(Rs!fecAlbar) & "," & db.fechahora(Rs!fecAlbar & " " & Format(Rs!horalbar, "hh:mm:ss")) & "," & db.numero(Rs!codTurno) & "," & db.numero(Rs!Numtarje) & "," & db.numero(Rs!codartic) & "," & _
              db.numero(Rs!cantidad) & "," & db.numero(Rs!preciove) & "," & db.numero(Rs!importel) & "," & db.Texto(Rs!matricul) & "," & _
              db.numero(Rs!Kilometros) & "," & db.numero(Rs!dtoalvic) & "," & db.numero(Rs!ImporteVale) & ")"
              
    End If
     
           
    InsertaLineaFactura = db.ejecutar(SQL)

eInsertaLineaFactura:
    If Err.Number <> 0 Or InsertaLineaFactura <> 0 Then
        MensError = "Error en la inserci�n de la l�nea de factura en el albaran " & Rs!numalbar
        If InsertaLineaFactura = 0 Then InsertaLineaFactura = 1
    End If
    
End Function

' en la facturacion ajena hemos de insertar en la temporal para luego hacer la factura global
Public Function InsertaLineaFacturaTemporal(ByRef db As BaseDatos, codartic As String, cantidad As String, importel As String) As Long
' importe1 = codartic
' importe2 = cantidad
' importe3 = importel

    Dim SQL As String
    Dim ImpLinea As Currency
    
    On Error GoTo eInsertaLineaFacturaTemporal
    MensError = ""
    
    SQL = "select count(*) from tmpinformes where importe1 = " & db.numero(codartic) & " and codusu = " & vSesion.Codigo
    
    If TotalRegistros(SQL) <> 0 Then
        SQL = "update tmpinformes set importe2 = importe2 + " & db.numero(cantidad) & ","
        SQL = SQL & "importe3 = importe3 + " & db.numero(importel)
        SQL = SQL & " where codusu = " & vSesion.Codigo & " and importe1 = " & db.numero(codartic)
    Else
        SQL = "insert into tmpinformes (codusu, importe1, importe2, importe3) values ("
        SQL = SQL & vSesion.Codigo & "," & db.numero(codartic) & "," & db.numero(cantidad) & ","
        SQL = SQL & db.numero(importel) & ")"
        
    End If
           
    InsertaLineaFacturaTemporal = db.ejecutar(SQL)

eInsertaLineaFacturaTemporal:
    If Err.Number <> 0 Or InsertaLineaFacturaTemporal <> 0 Then
        MensError = "Error en la inserci�n en temporal de la l�nea de albaran " & Rs!numalbar
        If InsertaLineaFacturaTemporal = 0 Then InsertaLineaFacturaTemporal = 1
    End If
    
End Function



Public Function InsertaLineaDescuento(ByRef db As BaseDatos, numser As String, numFac As Long, fecFac As Date, Linea As Integer, cantidad As Currency, Importe As Currency, Turno As Integer, Precio As Currency, Tarjeta As String, Tipo As Byte) As Long
' tipo = 0 facturacion normal
' tipo = 1 facturacion ajena
    Dim SQL As String
    Dim ImpLinea As Currency
    Dim Texto As String
    
    '05012007
    On Error GoTo eInsertaLineaDescuento
    MensError = ""
    Texto = "BONIFICA"
    
    If Tipo = 0 Then
        SQL = "INSERT into slhfac "
    Else
        SQL = "INSERT into slhfacr "
    End If
    
    SQL = SQL & "(letraser, numfactu, fecfactu, numlinea, numalbar, " & _
           "fecalbar, horalbar, codturno, numtarje, codartic, " & _
           "cantidad, preciove, implinea) " & _
           "values " & _
           "(" & db.Texto(numser) & "," & db.numero(numFac) & "," & db.Fecha(fecFac) & "," & db.numero(Linea) & "," & db.Texto(Texto) & "," & _
           db.Fecha(fecFac) & "," & db.fechahora(fecFac & " 0:00:00") & "," & db.numero(Turno) & "," & db.numero(Tarjeta) & "," & db.numero(vParamAplic.ArticDto) & "," & _
           db.numero(cantidad) & "," & db.numero(Precio) & "," & db.numero(Importe) & ")"
           
    InsertaLineaDescuento = db.ejecutar(SQL)
    
'05012007
eInsertaLineaDescuento:
    If Err.Number <> 0 Or InsertaLineaDescuento <> 0 Then
        MensError = "Error insertando en lineas de hist�rico de facturas una linea de descuento"
        If InsertaLineaDescuento = 0 Then InsertaLineaDescuento = 1
    End If
    
End Function

Public Function InsertaLineaDescuentoTemporal(ByRef db As BaseDatos, cantidad As Currency, Importe As Currency) As Long
    Dim SQL As String
    Dim ImpLinea As Currency
    Dim Texto As String
    
    On Error GoTo eInsertaLineaDescuentoTemporal
    MensError = ""
    
    SQL = "select count(*) from tmpinformes where importe1 = " & db.numero(vParamAplic.ArticDto) & " and codusu = " & vSesion.Codigo
    
    If TotalRegistros(SQL) <> 0 Then
        SQL = "update tmpinformes set importe2 = importe2 + " & db.numero(cantidad) & ","
        SQL = SQL & "importe3 = importe3 + " & db.numero(Importe)
        SQL = SQL & " where codusu = " & vSesion.Codigo & " and importe1 = " & db.numero(vParamAplic.ArticDto)
    Else
        SQL = "insert into tmpinformes (codusu, importe1, importe2, importe3) values ("
        SQL = SQL & vSesion.Codigo & "," & db.numero(vParamAplic.ArticDto) & "," & db.numero(cantidad) & ","
        SQL = SQL & db.numero(Importe) & ")"
        
    End If
           
    InsertaLineaDescuentoTemporal = db.ejecutar(SQL)
    
eInsertaLineaDescuentoTemporal:
    If Err.Number <> 0 Or InsertaLineaDescuentoTemporal <> 0 Then
        MensError = "Error insertando en temporal una linea de descuento"
        If InsertaLineaDescuentoTemporal = 0 Then InsertaLineaDescuentoTemporal = 1
    End If
    
End Function

Public Function Facturacion(db As BaseDatos, DesFec As String, HasFec As String, dessoc As String, hassoc As String, descop As String, hascop As String, FecFactura As Date, CliTar As Byte, Pb1 As ProgressBar, TipoClien As String, TipoGasoB As Byte, Optional TipoArt As Integer, Optional EsContado As Boolean) As Long
Dim SQL As String
Dim Rs As ADODB.Recordset

Dim Impuesto As Currency
Dim impbase As Currency
Dim ActSocio As Long
Dim ActForpa As Integer
Dim ActTarje As String
Dim AntAlbaran As String
Dim AntTarje As String
Dim AntSocio As Long
Dim AntForpa As Integer
Dim AntTurno As Integer

Dim AntDepar As Integer
Dim ActDepar As Integer

Dim HayReg As Boolean
Dim v_linea As Integer
Dim FamArtDto As String
Dim IvaArtDto As String
Dim ImporDto As Currency
Dim vCont As CContador
Dim DtoLitro As Currency
Dim CantCombustible As Currency
Dim Codigo As String
Dim baseimpo As Dictionary

Dim NumError As Long

Dim tipoMov As String



    On Error GoTo eFacturacion

    FamArtDto = "codfamia"
    IvaArtDto = DevuelveDesdeBD("codigiva", "sartic", "codartic", vParamAplic.ArticDto, "N", FamArtDto)
    
    SQL = "select scaalb.codclave, scaalb.codsocio, scaalb.codartic, scaalb.cantidad, scaalb.numlinea,"
    SQL = SQL & " scaalb.preciove, scaalb.importel, scaalb.numalbar, scaalb.fecalbar,"
    SQL = SQL & " scaalb.horalbar, scaalb.codturno, scaalb.codforpa, scaalb.numtarje, scaalb.matricul, scaalb.precioinicial, "
    SQL = SQL & " scaalb.kilometros,  "
    '[Monica]24/08/2015: a�adimos el dto alvic para el regaixo
    SQL = SQL & " scaalb.dtoalvic, "
    '[Monica]28/12/2015: a�adimos el importe de vale para el regaixo
    SQL = SQL & " scaalb.importevale "
    
    '[Monica]03/05/2019: a�adimos el departamento
    If CliTar = 4 Then
        SQL = SQL & ", starje.coddepar "
    End If
    
    '[Monica]03/05/2019: para el caso de facturar por departamentos
    If CliTar = 4 Then
        SQL = SQL & " from (((scaalb inner join ssocio on scaalb.codsocio = ssocio.codsocio) "
        SQL = SQL & " inner join scoope on ssocio.codcoope = scoope.codcoope "
        If descop <> "" Then SQL = SQL & " and ssocio.codcoope >= " & DBSet(descop, "N")
        If hascop <> "" Then SQL = SQL & " and ssocio.codcoope <= " & DBSet(hascop, "N")
    Else
        SQL = SQL & " from ((scaalb inner join ssocio on scaalb.codsocio = ssocio.codsocio) "
        SQL = SQL & " inner join scoope on ssocio.codcoope = scoope.codcoope "
        If descop <> "" Then SQL = SQL & " and ssocio.codcoope >= " & DBSet(descop, "N")
        If hascop <> "" Then SQL = SQL & " and ssocio.codcoope <= " & DBSet(hascop, "N")
    End If
    
    '[Monica]19/06/2013: A�adimos el if de cooperativa y tipogasob
    If (vParamAplic.Cooperativa = 1 Or vParamAplic.Cooperativa = 2) And TipoGasoB > 0 Then
        '[Monica]15/07/2013: a�adido el caso de que sea interna
        If CliTar = 3 Then
            SQL = SQL & " and scoope.tipfactu = " & DBLet(CliTar, "N") & ") "
        Else
            ' no miramos si es por cliente o por tarjeta
'            Sql = Sql & " and scoope.tipfactu <= " & DBLet(CliTar, "N") & ") "
            SQL = SQL & " and scoope.tipfactu in (0,1,4)) "
        End If
    Else
        SQL = SQL & " and scoope.tipfactu = " & DBLet(CliTar, "N") & ") "
    End If
    
    '[Monica]03/05/2019: facturacion por departamento
    If CliTar = 4 Then
        SQL = SQL & " inner join starje on scaalb.numtarje = starje.numtarje and scaalb.codsocio = starje.codsocio) "
    End If
    
    If vParamAplic.Cooperativa = 4 Then
        '[Monica]30/06/2014: antes solo se facturaba para pobla los articulos no combustibles (resto de productos)
        Select Case TipoArt
            Case 0 ' resto de productos
                SQL = SQL & " inner join sartic on scaalb.codartic = sartic.codartic and sartic.tipogaso = 0 "
            Case 1 ' gasolinas
                SQL = SQL & " inner join sartic on scaalb.codartic = sartic.codartic and sartic.tipogaso in (1,2,4) "
            Case 2 ' gasoleo B
                SQL = SQL & " inner join sartic on scaalb.codartic = sartic.codartic and sartic.tipogaso = 3 "
        End Select
    End If
    
    SQL = SQL & " where scaalb.numfactu = 0 and scaalb.codforpa <> 98 "
    If DesFec <> "" Then SQL = SQL & " and scaalb.fecalbar >= '" & Format(CDate(DesFec), FormatoFecha) & "' "
    If HasFec <> "" Then SQL = SQL & " and scaalb.fecalbar <= '" & Format(CDate(HasFec), FormatoFecha) & "' "
    If dessoc <> "" Then SQL = SQL & " and scaalb.codsocio >= " & DBSet(dessoc, "N")
    If hassoc <> "" Then SQL = SQL & " and scaalb.codsocio <= " & DBSet(hassoc, "N")
    
    '[Monica]29/12/2016: solo en el caso de ribarroja si es contado seleccionamos los de contado
    If vParamAplic.Cooperativa = 5 Then
        If EsContado Then
            SQL = SQL & " and scaalb.codforpa in (select codforpa from sforpa where tipforpa in (0,6))"
        Else
            SQL = SQL & " and scaalb.codforpa in (select codforpa from sforpa where not tipforpa in (0,6))"
        End If
    End If
    
    '[Monica]07/03/2018: solo los articulos que se facturan
    SQL = SQL & " and scaalb.codartic in (select codartic from sartic where facturar = 1) "
    
    
    Select Case TipoClien
        Case "0"
        
        Case "1"
            SQL = SQL & " and ssocio.bonifesp = 1"
        Case "2"
            SQL = SQL & " and ssocio.bonifesp = 0"
    End Select
    
    If vParamAplic.Cooperativa = 1 Or vParamAplic.Cooperativa = 2 Then
        Select Case TipoGasoB
            Case 0
                SQL = SQL & " and not scaalb.codartic in (select codartic from sartic where tipogaso = 3 union " & _
                                                         "select if(artdto is null, -1, artdto) from sartic where tipogaso = 3)"
            Case 1
                SQL = SQL & " and scaalb.codartic in (select codartic from sartic where tipogaso = 3 And esdomiciliado = 0 union " & _
                                                     "select if(artdto is null, -1, artdto) from sartic where tipogaso = 3 And esdomiciliado = 0)"
            Case 2
                SQL = SQL & " and scaalb.codartic in (select codartic from sartic where tipogaso = 3 And esdomiciliado = 1 union " & _
                                                     "select if(artdto is null, -1, artdto) from sartic where tipogaso = 3 And esdomiciliado = 1)"
        End Select
    End If
    
    '[Monica]28/07/2011: en el caso de las facturas internas quieren que sea por tarjeta antes era por cliente
    'If CliTar = 1 Or CliTar = 3 Then
    If CliTar = 1 Then
        SQL = SQL & " order by scaalb.codsocio, scaalb.codforpa, scaalb.fecalbar, scaalb.numalbar, scaalb.numlinea, scaalb.codclave "
    Else
        '[Monica]03/05/2019: para el caso de romper por departamento (que no sea gasoleo bonificado que es por tarjeta)
        If CliTar = 4 And TipoGasoB = 0 Then
            SQL = SQL & " order by scaalb.codsocio, starje.coddepar, scaalb.numtarje, scaalb.codforpa, scaalb.fecalbar, scaalb.numalbar, scaalb.numlinea, scaalb.codclave "
        Else
            SQL = SQL & " order by scaalb.codsocio, scaalb.numtarje, scaalb.codforpa, scaalb.fecalbar, scaalb.numalbar, scaalb.numlinea, scaalb.codclave "
        End If
    End If
    
    If CliTar = 3 Then
        '[Monica]15/07/2013: a�adida la condicion de tipo de gasoleo (nuevo tipo de movimiento para las internas gasoleo bonificado)
        Select Case TipoGasoB
            Case 0
                tipoMov = "FAI"
            Case 1, 2
                tipoMov = "FIB"
        End Select
    Else
        Select Case TipoGasoB
            Case 0
                tipoMov = "FAG"
                '[Monica]30/06/2014: para el caso de pobla del duc ya no hay facturacion cepsa y hay tres contadores
                If vParamAplic.Cooperativa = 4 Then
                    If TipoArt = 0 Then tipoMov = "FAG" ' facturas de resto de productos
                    If TipoArt = 1 Then tipoMov = "FGA" ' facturas de gasolina
                    If TipoArt = 2 Then tipoMov = "FGB" ' facturas de gasoleo B
                End If
                
                '[Monica]29/12/2016: para el caso de Ribarroja hay 2 tipos de contadores (uno para contados y otro no)
                If vParamAplic.Cooperativa = 5 Then
                    If EsContado Then tipoMov = "FA1"
                End If
                
            Case 1 'Gasoleo B
                tipoMov = "FGB"
            Case 2 'Gasoleo B Domiciliado
                tipoMov = "FGD"
        End Select
    End If
    
    Set Rs = db.cursor(SQL)
    HayReg = False
    v_linea = 0
    NumError = 0
    If Not Rs.EOF Then
        Rs.MoveFirst
        AntSocio = Rs!codsocio
        AntAlbaran = Rs!numalbar
        AntForpa = Rs!Codforpa
        AntTurno = Rs!codTurno
        AntTarje = Rs!Numtarje
        If CliTar = 4 Then AntDepar = DBLet(Rs!coddepar, "N")
        
        Set baseimpo = New Dictionary
        ' cogemos el numero de factura de parametros
        
        Set vCont = New CContador
        If Not vCont.ConseguirContador(tipoMov, True, db) Then Exit Function
        
        numser = ""
        numser = DevuelveDesdeBD("letraser", "stipom", "codtipom", tipoMov, "T")
        
        TotalImp = 0
        TotalImpSigaus = 0
        ImpFactu = 0
        
        While Not Rs.EOF And NumError = 0
            HayReg = True
            ActForpa = Rs!Codforpa
            ActSocio = Rs!codsocio
            ActTarje = Rs!Numtarje
            '[Monica]03/05/2019: facturacion por departamentos
            If CliTar = 4 Then ActDepar = DBLet(Rs!coddepar, "N")
            '[Monica]23/07/2013     ' after group of codforpa
            If ((ActForpa <> AntForpa Or ActSocio <> AntSocio) And (CliTar = 1 Or (CliTar = 3 And TipoGasoB = 0))) Or _
            ((ActForpa <> AntForpa Or ActSocio <> AntSocio Or ActTarje <> AntTarje) And (CliTar = 0 Or (CliTar = 3 And TipoGasoB <> 0))) Or _
            ((ActForpa <> AntForpa Or ActSocio <> AntSocio Or ActDepar <> AntDepar) And CliTar = 4) Then
            
               '  ### [Monica] 05/12/2006
               ' modificacion: si la forma de pago no admite bonificacion no hacemos
               If AdmiteBonificacion(AntForpa) Then
 
                   ' miramos el descuento/litro de socio sobre la cantidad
                   SQL = ""
                   SQL = DevuelveDesdeBD("dtolitro", "ssocio", "codsocio", CStr(AntSocio), "N")
                   DtoLitro = 0
                   If SQL <> "" Then DtoLitro = CCur(SQL)
    
                   If DtoLitro <> 0 Then
                        DtoLitro = DtoLitro * (-1)
                        ImporDto = Round2(CantCombustible * DtoLitro, 2)
                        baseimpo(Val(IvaArtDto)) = DBLet(baseimpo(Val(IvaArtDto)), "N") + DBLet(ImporDto, "N")
                        v_linea = v_linea + 1
                        If NumError = 0 Then NumError = InsertaLineaDescuento(db, numser, vCont.Contador, FecFactura, v_linea, CantCombustible, ImporDto, AntTurno, DtoLitro, AntTarje, 0)
                   End If
               End If
               
               v_linea = 0
               
               If NumError = 0 Then
                    If CliTar = 3 Then
                        NumError = InsertCabe(db, baseimpo, vCont.Contador, FecFactura, AntSocio, AntForpa, 0, False, True)
                    Else
                        NumError = InsertCabe(db, baseimpo, vCont.Contador, FecFactura, AntSocio, AntForpa, 0, False, False, AntDepar)
                    End If
               End If

               '[Monica]01/08/2011: Insertamos solo en la svenci en la facturacion ya que la insercion en tesoreria
               '                    se hace en la contabilizacion de facturas dada una fecha de vencimiento
               If NumError = 0 Then
                    TipForpa = DevuelveDesdeBDNew(cPTours, "sforpa", "tipforpa", "codforpa", CStr(AntForpa), "N")
                    '[Monica]04/01/2013: efectivos
                    If TipForpa <> "0" And TipForpa <> "6" Then
                        NumError = InsertarVencimientos(db, numser, vCont.Contador, CStr(FecFactura), CStr(AntForpa))
                    End If
               End If
               
               Set baseimpo = Nothing
               Set baseimpo = New Dictionary
               TotalImp = 0
               TotalImpSigaus = 0
               AntForpa = ActForpa
               AntSocio = ActSocio
               AntTurno = Rs!codTurno
               AntTarje = ActTarje
               '[Monica]03/05/2019: departamentos
               AntDepar = ActDepar
               
               CantCombustible = 0
            
                '[Monica]24/01/2013: si el socio es un cliente no de varios vemos si hay q partirle la factura
               ImpFactu = 0
               
               If Not vCont.ConseguirContador(tipoMov, True, db) Then Exit Function
            End If
            
            '[Monica]24/01/2013: si el socio es un cliente no de varios vemos si hay q partirle la factura
            TipForpa = DevuelveDesdeBDNew(cPTours, "sforpa", "tipforpa", "codforpa", CStr(AntForpa), "N")
            If vParamAplic.Cooperativa = 1 And Not EsDeVarios(CStr(AntSocio)) And vParamAplic.LimiteFra <> 0 And (ImpFactu + DBLet(Rs!importel, "N") > vParamAplic.LimiteFra) And TipForpa = "0" Then
           
               If NumError = 0 Then
                    If CliTar = 3 Then
                        NumError = InsertCabe(db, baseimpo, vCont.Contador, FecFactura, AntSocio, AntForpa, 0, False, True)
                    Else
                        NumError = InsertCabe(db, baseimpo, vCont.Contador, FecFactura, AntSocio, AntForpa, 0, False, False, AntDepar)
                    End If
               End If

               '[Monica]01/08/2011: Insertamos solo en la svenci en la facturacion ya que la insercion en tesoreria
               '                    se hace en la contabilizacion de facturas dada una fecha de vencimiento
               If NumError = 0 Then
                    TipForpa = DevuelveDesdeBDNew(cPTours, "sforpa", "tipforpa", "codforpa", CStr(AntForpa), "N")
                    '[Monica]04/01/2013: efectivos
                    If TipForpa <> "0" And TipForpa <> "6" Then
                        NumError = InsertarVencimientos(db, numser, vCont.Contador, CStr(FecFactura), CStr(AntForpa))
                    End If
               End If
               
               Set baseimpo = Nothing
               Set baseimpo = New Dictionary
               TotalImp = 0
               TotalImpSigaus = 0
               
               CantCombustible = 0
               
               ImpFactu = 0
               
               If Not vCont.ConseguirContador(tipoMov, True, db) Then Exit Function
           
            Else
                '[Monica]24/01/2013: a�ado esta variable de importe total de factura para ver si se pasa de la cantidad de parametros
                ImpFactu = ImpFactu + DBLet(Rs!importel, "N")
                
                '-------
                ' tenemos que calcular el impuesto multiplicando cantidad de linea por impuesto por articulo
                Codigo = "codigiva"
                Sql1 = ""
                Sql1 = DevuelveDesdeBD("impuesto", "sartic", "codartic", DBLet(Rs!codartic), "N", Codigo)
                If Sql1 = "" Then
                    Impuesto = 0
                Else
                    Impuesto = CCur(Sql1) ' Comprueba si es nulo y lo pone a 0 o ""
                End If
                
                If EsArticuloCombustible(Rs!codartic) Then
                    TotalImp = TotalImp + Round2((Rs!cantidad * Impuesto), 2)
                    CantCombustible = CantCombustible + DBLet(Rs!cantidad, "N")
                End If
                
                '[Monica]15/02/2011: cuando el articulo es lubricante, tiene un impuesto, hemos de calcularlo
                ' Sabemos que es lubricante pq tiene un peso por unidad.
                ' El Impuesto se calcula multiplicandolo por el preciosigaus
                TotalImpSigaus = TotalImpSigaus + ImpuestoSigausArticulo(Rs!codartic, Rs!cantidad)
                
                
                baseimpo(Val(Codigo)) = DBLet(baseimpo(Val(Codigo)), "N") + DBLet(Rs!importel, "N")
                v_linea = v_linea + 1
                
                IncrementarProgres Pb1, 1
                
                If NumError = 0 Then NumError = InsertaLineaFactura(db, Rs, numser, vCont.Contador, FecFactura, v_linea, 0)
                If NumError = 0 Then NumError = BorrarLineaAlbaran(db, Rs!Codclave, True)
    
                'Siguiente
        '        antfactura = Rs!numfactu
                'If CliTar = 1 Then AntTarje = ActTarje (RAFA)
        
                Rs.MoveNext
            
            End If
        Wend
        If HayReg And NumError = 0 Then
               
               ' miramos el descuento/litro de socio sobre la cantidad
               
               If AdmiteBonificacion(AntForpa) Then
                    SQL = ""
                    SQL = DevuelveDesdeBD("dtolitro", "ssocio", "codsocio", CStr(AntSocio), "N")
                    DtoLitro = 0
                    If SQL <> "" Then DtoLitro = CCur(SQL)
                    If DtoLitro <> 0 Then
                         DtoLitro = DtoLitro * (-1)
                         ImporDto = Round2(CantCombustible * DtoLitro, 2)
                         baseimpo(Val(IvaArtDto)) = DBLet(baseimpo(Val(IvaArtDto)), "N") + DBLet(ImporDto, "N")
                         v_linea = v_linea + 1
                         If NumError = 0 Then NumError = InsertaLineaDescuento(db, numser, vCont.Contador, FecFactura, v_linea, CantCombustible, ImporDto, AntTurno, DtoLitro, AntTarje, 0)
                    End If
               End If
               If NumError = 0 Then
                    If CliTar = 3 Then
                        NumError = InsertCabe(db, baseimpo, vCont.Contador, FecFactura, AntSocio, AntForpa, 0, False, True)
                    Else
                        NumError = InsertCabe(db, baseimpo, vCont.Contador, FecFactura, AntSocio, AntForpa, 0, False, False, AntDepar)
                    End If
               End If
               
               '[Monica]01/08/2011: Insertamos solo en la svenci en la facturacion ya que la insercion en tesoreria
               '                    se hace en la contabilizacion de facturas dada una fecha de vencimiento
               If NumError = 0 Then
                    TipForpa = DevuelveDesdeBDNew(cPTours, "sforpa", "tipforpa", "codforpa", CStr(AntForpa), "N")
                    '[Monica]04/01/2013: efectivos
                    If TipForpa <> "0" And TipForpa <> "6" Then
                        NumError = InsertarVencimientos(db, numser, vCont.Contador, CStr(FecFactura), CStr(AntForpa))
                    End If
               End If
               
               If NumError = 0 Then
               
                   '[Monica]07/03/2018: borramos los albaranes que son articulos que no se facturan
                    Dim SqlNue As String
                    
                    SqlNue = "delete from scaalb where codsocio  in (select codsocio from ssocio inner join scoope on ssocio.codcoope = scoope.codcoope where (1=1) "
                    If descop <> "" Then SqlNue = SqlNue & " and ssocio.codcoope >= " & DBSet(descop, "N")
                    If hascop <> "" Then SqlNue = SqlNue & " and ssocio.codcoope <= " & DBSet(hascop, "N")
                    
                    '[Monica]19/06/2013: A�adimos el if de cooperativa y tipogasob
                    If (vParamAplic.Cooperativa = 1 Or vParamAplic.Cooperativa = 2) And TipoGasoB > 0 Then
                        '[Monica]15/07/2013: a�adido el caso de que sea interna
                        If CliTar = 3 Then
                            SqlNue = SqlNue & " and scoope.tipfactu = " & DBLet(CliTar, "N")
                        Else
                            ' no miramos si es por cliente o por tarjeta
                '            Sql = Sql & " and scoope.tipfactu <= " & DBLet(CliTar, "N") & ") "
                            SqlNue = SqlNue & " and scoope.tipfactu in (0,1) "
                        End If
                    Else
                        SqlNue = SqlNue & " and scoope.tipfactu = " & DBLet(CliTar, "N")
                    End If
                    
                    Select Case TipoClien
                        Case "0"
                        
                        Case "1"
                            SqlNue = SqlNue & " and ssocio.bonifesp = 1"
                        Case "2"
                            SqlNue = SqlNue & " and ssocio.bonifesp = 0"
                    End Select
                    
                    
                    SqlNue = SqlNue & ") "
                    
                    SqlNue = SqlNue & " and scaalb.codartic in (select codartic from sartic where (1=1) "
                    
                    If vParamAplic.Cooperativa = 4 Then
                        '[Monica]30/06/2014: antes solo se facturaba para pobla los articulos no combustibles (resto de productos)
                        Select Case TipoArt
                            Case 0 ' resto de productos
                                SqlNue = SqlNue & "  and sartic.tipogaso = 0 "
                            Case 1 ' gasolinas
                                SqlNue = SqlNue & "  and sartic.tipogaso in (1,2,4) "
                            Case 2 ' gasoleo B
                                SqlNue = SqlNue & "  and sartic.tipogaso = 3 "
                        End Select
                    End If
                    '[Monica]07/03/2018: solo los articulos que no se facturan
                    SqlNue = SqlNue & " and facturar = 0) and scaalb.numfactu = 0 and scaalb.codforpa <> 98 "
                    
                    If DesFec <> "" Then SqlNue = SqlNue & " and scaalb.fecalbar >= '" & Format(CDate(DesFec), FormatoFecha) & "' "
                    If HasFec <> "" Then SqlNue = SqlNue & " and scaalb.fecalbar <= '" & Format(CDate(HasFec), FormatoFecha) & "' "
                    If dessoc <> "" Then SqlNue = SqlNue & " and scaalb.codsocio >= " & DBSet(dessoc, "N")
                    If hassoc <> "" Then SqlNue = SqlNue & " and scaalb.codsocio <= " & DBSet(hassoc, "N")
                    
                    '[Monica]29/12/2016: solo en el caso de ribarroja si es contado seleccionamos los de contado
                    If vParamAplic.Cooperativa = 5 Then
                        If EsContado Then
                            SqlNue = SqlNue & " and scaalb.codforpa in (select codforpa from sforpa where tipforpa in (0,6))"
                        Else
                            SqlNue = SqlNue & " and scaalb.codforpa in (select codforpa from sforpa where not tipforpa in (0,6))"
                        End If
                    End If
                    
                    
                    If vParamAplic.Cooperativa = 1 Or vParamAplic.Cooperativa = 2 Then
                        Select Case TipoGasoB
                            Case 0
                                SqlNue = SqlNue & " and not scaalb.codartic in (select codartic from sartic where tipogaso = 3 union " & _
                                                                         "select if(artdto is null, -1, artdto) from sartic where tipogaso = 3)"
                            Case 1
                                SqlNue = SqlNue & " and scaalb.codartic in (select codartic from sartic where tipogaso = 3 And esdomiciliado = 0 union " & _
                                                                     "select if(artdto is null, -1, artdto) from sartic where tipogaso = 3 And esdomiciliado = 0)"
                            Case 2
                                SqlNue = SqlNue & " and scaalb.codartic in (select codartic from sartic where tipogaso = 3 And esdomiciliado = 1 union " & _
                                                                     "select if(artdto is null, -1, artdto) from sartic where tipogaso = 3 And esdomiciliado = 1)"
                        End Select
                    End If
                               
                                               
                    NumError = BorrarAlbaranes(db, SqlNue)
                           
                           
               End If
               
               
               
               
        End If
    End If
eFacturacion:
    Facturacion = NumError
    Exit Function
End Function

Private Function BorrarAlbaranes(ByRef db As BaseDatos, SQL As String) As Long

    On Error GoTo eBorrarAlbaranes

    NumError = db.ejecutar(SQL)
    
eBorrarAlbaranes:
    If Err.Number <> 0 Then NumError = Err.Number
    
    BorrarAlbaranes = NumError
    Exit Function
End Function



Private Function InsertarVencimientos(ByRef db As BaseDatos, Serie As String, Factura As String, FecFactura As String, ForPago As String) As Long
Dim SQL As String
Dim Rs As ADODB.Recordset
Dim RsTotal As ADODB.Recordset
Dim rsVenci As ADODB.Recordset
Dim SQLinsert As String
Dim SqlValues As String
Dim SqlAux As String
Dim ImpVenci As Currency
Dim FecVenci As Date
Dim I As Integer
Dim TotalFac As Currency
    On Error GoTo eInsertarVencimientos
    
    InsertarVencimientos = 0


    SQLinsert = "insert into svenci (letraser, numfactu, fecfactu, ordefect, fecefect, impefect) values "

    SqlAux = DBSet(Serie, "T") & "," & DBSet(Factura, "N") & "," & DBSet(FecFactura, "F")

    SQL = "select totalfac from schfac where letraser = " & DBSet(Serie, "T") & " and numfactu = " & DBSet(Factura, "N") & " and fecfactu = " & DBSet(FecFactura, "F")
    Set RsTotal = db.cursor(SQL)
    TotalFac = DBLet(RsTotal.Fields(0).Value, "N")
    Set RsTotal = Nothing
    
    'Obtener el N� de Vencimientos de la forma de pago
    SQL = "SELECT numerove, diasvto primerve, restoven FROM sforpa WHERE codforpa=" & ForPago

    Set rsVenci = db.cursor(SQL)
    
    If Not rsVenci.EOF Then
        If rsVenci!numerove > 0 And CCur(TotalFac) <> 0 Then
        
            '-------- Primer Vencimiento
            I = 1
            'FECHA VTO
            FecVenci = CDate(FecFactura)
            FecVenci = DateAdd("d", DBLet(rsVenci!primerve, "N"), FecVenci)
            '===
        
            'IMPORTE del Vencimiento
            TotalFactura2 = TotalFac
            If rsVenci!numerove = 1 Then
                ImpVenci = TotalFactura2
            Else
                ImpVenci = Round2(TotalFactura2 / rsVenci!numerove, 2)
                'Comprobar que la suma de los vencimientos cuadra con el total de la factura
                If ImpVenci * rsVenci!numerove <> TotalFactura2 Then
                    ImpVenci = Round(ImpVenci + (TotalFactura2 - ImpVenci * rsVenci!numerove), 2)
                End If
            End If

            SqlValues = "(" & SqlAux & "," & DBSet(I, "N") & "," & DBSet(FecVenci, "F") & "," & DBSet(ImpVenci, "N") & "),"

            'Resto Vencimientos
            '--------------------------------------------------------------------
            For I = 2 To rsVenci!numerove
               'FECHA Resto Vencimientos
                FecVenci = DateAdd("d", DBLet(rsVenci!restoven, "N"), FecVenci)
                '===
                
                'IMPORTE Resto de Vendimientos
                ImpVenci = Round2(TotalFactura2 / rsVenci!numerove, 2)
                
                SqlValues = SqlValues & "(" & SqlAux & "," & DBSet(I, "N") & "," & DBSet(FecVenci, "F") & "," & DBSet(ImpVenci, "N") & "),"
            Next I

            If SqlValues <> "" Then
                SqlValues = Mid(SqlValues, 1, Len(SqlValues) - 1)
                
                InsertarVencimientos = db.ejecutar(SQLinsert & SqlValues)
            End If
        End If
    End If
    Set rsVenci = Nothing
    
    Exit Function

eInsertarVencimientos:
    InsertarVencimientos = Err.Number
End Function





Public Function FacturacionCepsa(db As BaseDatos, DesFec As String, HasFec As String, dessoc As String, hassoc As String, descop As String, hascop As String, FecFactura As Date, CliTar As Byte, Pb1 As ProgressBar, FecVenci As String, Banpr As String) As Long
Dim SQL As String
Dim Rs As ADODB.Recordset

Dim Impuesto As Currency
Dim impbase As Currency
Dim ActSocio As Long
Dim ActForpa As Integer
Dim ActTarje As String 'Long
Dim AntAlbaran As String
Dim AntTarje As String 'Long
Dim AntSocio As Long
Dim AntForpa As Integer
Dim AntTurno As Integer
Dim HayReg As Boolean
Dim v_linea As Integer
Dim FamArtDto As String
Dim IvaArtDto As String
Dim ImporDto As Currency
Dim vCont As CContador
Dim DtoLitro As Currency
Dim CantCombustible As Currency
Dim Codigo As String
Dim baseimpo As Dictionary
Dim vWhere1 As String

Dim NumError As Long
Dim MenError As String
Dim vsocio As CSocio

Dim TipForpa As String


    On Error GoTo eFacturacion

    FamArtDto = "codfamia"
    IvaArtDto = DevuelveDesdeBD("codigiva", "sartic", "codartic", vParamAplic.ArticDto, "N", FamArtDto)
    
    SQL = "select scaalb.codclave, scaalb.codsocio, scaalb.codartic, scaalb.cantidad, scaalb.numlinea,"
    SQL = SQL & " scaalb.preciove, scaalb.importel, scaalb.numalbar, scaalb.fecalbar,"
    SQL = SQL & " scaalb.horalbar, scaalb.codturno, scaalb.codforpa, scaalb.numtarje, scaalb.matricul, scaalb.precioinicial, scaalb.kilometros "
    SQL = SQL & " from ((scaalb inner join ssocio on scaalb.codsocio = ssocio.codsocio) "
    SQL = SQL & " inner join scoope on ssocio.codcoope = scoope.codcoope "
    If descop <> "" Then SQL = SQL & " and ssocio.codcoope >= " & DBSet(descop, "N")
    If hascop <> "" Then SQL = SQL & " and ssocio.codcoope <= " & DBSet(hascop, "N")
    SQL = SQL & " and scoope.tipfactu = " & DBLet(CliTar, "N") & ") "
    
    SQL = SQL & " inner join sartic on scaalb.codartic = sartic.codartic and sartic.tipogaso <> 0    "
    
    SQL = SQL & " where scaalb.numfactu = 0 and scaalb.codforpa <> 98 "
    If DesFec <> "" Then SQL = SQL & " and scaalb.fecalbar >= '" & Format(CDate(DesFec), FormatoFecha) & "' "
    If HasFec <> "" Then SQL = SQL & " and scaalb.fecalbar <= '" & Format(CDate(HasFec), FormatoFecha) & "' "
    If dessoc <> "" Then SQL = SQL & " and scaalb.codsocio >= " & DBSet(dessoc, "N")
    If hassoc <> "" Then SQL = SQL & " and scaalb.codsocio <= " & DBSet(hassoc, "N")
    
    
    If CliTar = 1 Then
        SQL = SQL & " order by scaalb.codsocio, scaalb.codforpa, scaalb.fecalbar, scaalb.numalbar, scaalb.numlinea "
    Else
        SQL = SQL & " order by scaalb.codsocio, scaalb.numtarje, scaalb.codforpa, scaalb.fecalbar, scaalb.numalbar, scaalb.numlinea "
    End If
    
    Set Rs = db.cursor(SQL)
    HayReg = False
    v_linea = 0
    NumError = 0
    If Not Rs.EOF Then
        Rs.MoveFirst
        AntSocio = Rs!codsocio
        AntAlbaran = Rs!numalbar
        AntForpa = Rs!Codforpa
        AntTurno = Rs!codTurno
        AntTarje = Rs!Numtarje
        
        Set baseimpo = New Dictionary
        ' cogemos el numero de factura de parametros
        
        Set vCont = New CContador
        If Not vCont.ConseguirContador("FAC", True, db) Then Exit Function
        
        numser = ""
        numser = DevuelveDesdeBD("letraser", "stipom", "codtipom", "FAC", "T")
        
        TotalImp = 0
        
        While Not Rs.EOF And NumError = 0
            HayReg = True
            ActForpa = Rs!Codforpa
            ActSocio = Rs!codsocio
            ActTarje = Rs!Numtarje
            If ((ActForpa <> AntForpa Or ActSocio <> AntSocio) And CliTar = 1) Or _
            ((ActForpa <> AntForpa Or ActSocio <> AntSocio Or ActTarje <> AntTarje) And CliTar = 0) Then ' after group of codforpa
            
               '  ### [Monica] 05/12/2006
               ' modificacion: si la forma de pago no admite bonificacion no hacemos
               If AdmiteBonificacion(AntForpa) Then
 
                   ' miramos el descuento/litro de socio sobre la cantidad
                   SQL = ""
                   SQL = DevuelveDesdeBD("dtolitro", "ssocio", "codsocio", CStr(AntSocio), "N")
                   DtoLitro = 0
                   If SQL <> "" Then DtoLitro = CCur(SQL)
    
                   If DtoLitro <> 0 Then
                        DtoLitro = DtoLitro * (-1)
                        ImporDto = Round2(CantCombustible * DtoLitro, 2)
                        baseimpo(Val(IvaArtDto)) = DBLet(baseimpo(Val(IvaArtDto)), "N") + DBLet(ImporDto, "N")
                        v_linea = v_linea + 1
                        If NumError = 0 Then NumError = InsertaLineaDescuento(db, numser, vCont.Contador, FecFactura, v_linea, CantCombustible, ImporDto, AntTurno, DtoLitro, AntTarje, 0)
                   End If
               End If
               
               v_linea = 0
               
               If NumError = 0 Then NumError = InsertCabe(db, baseimpo, vCont.Contador, FecFactura, AntSocio, AntForpa, 0, True)
               
               If NumError = 0 Then
                    vWhere1 = "letraser = " & DBSet(numser, "T") & " and numfactu = " & DBSet(vCont.Contador, "N") & " and fecfactu = " & DBSet(FecFactura, "F")
                    MenError = "Insertando en tesoreria:"
                    Set vsocio = New CSocio
                    If vsocio.LeerDatos(CStr(AntSocio)) Then
'[Monica]16/12/2010: en Pobla se inserta todo en tesoreria pq no hay contabizacion de cierre de turno
'                        TipForpa = DevuelveDesdeBDNew(cPTours, "sforpa", "tipforpa", "codforpa", CStr(AntForpa), "N")
'                        If TipForpa <> "0" Then
                            If Not InsertarEnTesoreriaDB(db, vWhere1, FecVenci, Banpr, MenError, vsocio, "schfac") Then
                                NumError = 1
                                MsgBox MenError, vbExclamation
                            End If
'                        End If
                    Else
                        NumError = 1
                    End If
                    Set vsocio = Nothing
               End If
               
               Set baseimpo = Nothing
               Set baseimpo = New Dictionary
               TotalImp = 0
               AntForpa = ActForpa
               AntSocio = ActSocio
               AntTurno = Rs!codTurno
               AntTarje = ActTarje
               
               CantCombustible = 0
               
               If Not vCont.ConseguirContador("FAC", True, db) Then Exit Function
            End If
            
            '-------
            ' tenemos que calcular el impuesto multiplicando cantidad de linea por impuesto por articulo
            Codigo = "codigiva"
            Sql1 = ""
            Sql1 = DevuelveDesdeBD("impuesto", "sartic", "codartic", DBLet(Rs!codartic), "N", Codigo)
            If Sql1 = "" Then
                Impuesto = 0
            Else
                Impuesto = CCur(Sql1) ' Comprueba si es nulo y lo pone a 0 o ""
            End If
            
            If EsArticuloCombustible(Rs!codartic) Then
                TotalImp = TotalImp + Round2((Rs!cantidad * Impuesto), 2)
                CantCombustible = CantCombustible + DBLet(Rs!cantidad, "N")
            End If
            
            baseimpo(Val(Codigo)) = DBLet(baseimpo(Val(Codigo)), "N") + DBLet(Rs!importel, "N")
            v_linea = v_linea + 1
            
            IncrementarProgres Pb1, 1
            
            If NumError = 0 Then NumError = InsertaLineaFactura(db, Rs, numser, vCont.Contador, FecFactura, v_linea, 0)
            If NumError = 0 Then NumError = BorrarLineaAlbaran(db, Rs!Codclave, True)

            'Siguiente
    '        antfactura = Rs!numfactu
            'If CliTar = 1 Then AntTarje = ActTarje (RAFA)
    
            Rs.MoveNext
        Wend
        If HayReg And NumError = 0 Then
               ' miramos el descuento/litro de socio sobre la cantidad
               
               If AdmiteBonificacion(AntForpa) Then
                    SQL = ""
                    SQL = DevuelveDesdeBD("dtolitro", "ssocio", "codsocio", CStr(AntSocio), "N")
                    DtoLitro = 0
                    If SQL <> "" Then DtoLitro = CCur(SQL)
                    If DtoLitro <> 0 Then
                         DtoLitro = DtoLitro * (-1)
                         ImporDto = Round2(CantCombustible * DtoLitro, 2)
                         baseimpo(Val(IvaArtDto)) = DBLet(baseimpo(Val(IvaArtDto)), "N") + DBLet(ImporDto, "N")
                         v_linea = v_linea + 1
                         If NumError = 0 Then NumError = InsertaLineaDescuento(db, numser, vCont.Contador, FecFactura, v_linea, CantCombustible, ImporDto, AntTurno, DtoLitro, AntTarje, 0)
                    End If
               End If
               If NumError = 0 Then NumError = InsertCabe(db, baseimpo, vCont.Contador, FecFactura, AntSocio, AntForpa, 0, True)
               
               If NumError = 0 Then
                    vWhere1 = "letraser = " & DBSet(numser, "T") & " and numfactu = " & DBSet(vCont.Contador, "N") & " and fecfactu = " & DBSet(FecFactura, "F")
                    MenError = "Insertando en tesoreria:"
                    Set vsocio = New CSocio
                    If vsocio.LeerDatos(CStr(AntSocio)) Then
'[Monica]16/12/2010: en Pobla se inserta todo en tesoreria pq no hay contabizacion de cierre de turno
'                        TipForpa = DevuelveDesdeBDNew(cPTours, "sforpa", "tipforpa", "codforpa", CStr(AntForpa), "N")
'                        If TipForpa <> "0" Then
                            If Not InsertarEnTesoreriaDB(db, vWhere1, FecVenci, Banpr, MenError, vsocio, "schfac") Then
                                NumError = 1
                                MsgBox MenError, vbExclamation
                            End If
'                        End If
                    Else
                        NumError = 1
                    End If
                    Set vsocio = Nothing
               End If
               
        End If
    End If
eFacturacion:
    FacturacionCepsa = NumError
    Exit Function
End Function



Public Function BorrarLineaAlbaran(ByRef db As BaseDatos, clave As Long, DentroTransaccion As Boolean) As Long
Dim SQL As String

    SQL = "delete from scaalb where codclave = " & DBLet(clave, "N")
    
    BorrarLineaAlbaran = db.ejecutar(SQL)

End Function

Public Function BorradoAlbaranes(ByRef db As BaseDatos, desde As String, hasta As String) As Long
Dim Sql1 As String
    Sql1 = "delete from scaalb where numfactu <> 0 "
    
    If desde <> "" Then Sql1 = Sql1 & " and fecalbar >= '" & Format(desde, FormatoFecha) & "'"
    If hasta <> "" Then Sql1 = Sql1 & " and fecalbar <= '" & Format(hasta, FormatoFecha) & "'"
            
    BorradoAlbaranes = db.ejecutar(Sql1)
End Function

Public Function BorradoAlbaranesPrueba(ByRef db As BaseDatos, desde As String, hasta As String) As Long
Dim Sql1 As String

    Sql1 = "delete from scaalb where codforpa = 98 "
    
    If desde <> "" Then Sql1 = Sql1 & " and fecalbar >= '" & Format(desde, FormatoFecha) & "'"
    If hasta <> "" Then Sql1 = Sql1 & " and fecalbar <= '" & Format(hasta, FormatoFecha) & "'"
    
    BorradoAlbaranesPrueba = db.ejecutar(Sql1)
End Function
 
Public Function AdmiteBonificacion(Forpa As Integer) As Boolean
Dim SQL As String

    SQL = ""
    SQL = DevuelveDesdeBD("permitebonif", "sforpa", "codforpa", CStr(Forpa), "N")
    
    AdmiteBonificacion = (SQL = "1")

End Function


Public Function FechaSuperiorUltimaLiquidacion(fec As Date) As Boolean
Dim SQL As String
Dim Rs As ADODB.Recordset
Dim Mensual As Boolean
Dim Anofactu As Integer
Dim PeriodoFactu As Integer
Dim FechaDesde As Date

    On Error GoTo eFechaSuperiorUltimaLiquidacion

    FechaSuperiorUltimaLiquidacion = False

    ' en caso de que haya contabilidad comprobamos que la fecha de factura introducida
    ' no sea inferior a la ultima liquidacion de iva.
    If vParamAplic.NumeroConta <> 0 Then
        SQL = "select periodos, anofactu, perfactu from parametros"
        Set Rs = New ADODB.Recordset
        Rs.Open SQL, ConnConta, adOpenDynamic, adLockOptimistic
        
        If Not Rs.EOF Then
            Mensual = (Rs.Fields(0).Value = 1)
            Anofactu = Rs.Fields(1).Value
            PeriodoFactu = Rs.Fields(2).Value
            
            If Mensual Then ' facturacion mensual
                If PeriodoFactu = 12 Then
                    FechaDesde = CDate("01/01/" & Format(Anofactu + 1, "0000"))
                Else
                    FechaDesde = CDate("01/" & Format(PeriodoFactu + 1, "00") & "/" & Format(Anofactu, "0000"))
                End If
            Else ' facturacion trimestral
                If PeriodoFactu = 4 Then
                    FechaDesde = CDate("01/01/" & Format(Anofactu + 1, "0000"))
                Else
                    FechaDesde = CDate("01/" & Format((PeriodoFactu * 3) + 1, "00") & "/" & Format(Anofactu, "0000"))
                End If
            End If
            
            FechaSuperiorUltimaLiquidacion = (fec >= FechaDesde)
        End If
    End If

eFechaSuperiorUltimaLiquidacion:
    If Err.Number <> 0 Then
         MuestraError Err.Number, Err.Description
    End If
End Function


Public Function FechaDentroPeriodoContable(fec As Date) As Boolean
Dim SQL As String
Dim Rs As ADODB.Recordset
Dim Mensual As Boolean
Dim Anofactu As Integer
Dim PeriodoFactu As Integer
Dim FechaDesde As Date

    On Error GoTo eFechaDentroPeriodoContable

    FechaDentroPeriodoContable = (CDate(FIni) <= fec) And (fec <= (CDate(FFin) + 365))

eFechaDentroPeriodoContable:
    If Err.Number <> 0 Then
         MuestraError Err.Number, Err.Description
    End If
End Function

Public Function FechaFacturaInferiorUltimaFacturaSerieHco(Fecha As Date, numfactu As Long, Serie As String, Tipo As Byte) As Boolean
' tipo = 0 indica schfac
' tipo = 1 indica schfac2 hco.de ajenas del Regaixo
Dim SQL As String
Dim Rs As ADODB.Recordset

    FechaFacturaInferiorUltimaFacturaSerieHco = False

    SQL = "select fecfactu "
    If Tipo = 0 Then
        SQL = SQL & "from schfac "
    Else
        SQL = SQL & "from schfacr "
    End If
    SQL = SQL & " where numfactu = " & DBSet(numfactu, "N") & " and letraser = " & DBSet(Serie, "T")
    
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, Conn, adOpenDynamic, adLockOptimistic
    If Not Rs.EOF Then
        If Rs.Fields(0).Value > Fecha Then
            FechaFacturaInferiorUltimaFacturaSerieHco = True
        End If
    End If

End Function
'
' FACTURACION AJENA UTILIZADA EN EL REGAIXO PARA FACTURAR CLIENTES QUE SEAN SOCIOS DE LLOMBAI O CATADAU
' Se hace una factura por cada uno de los clientes y finalmente una factura global a Catadau o Llombai
' detallando totales por articulo.
'
Public Function FacturacionAjena(db As BaseDatos, DesFec As String, HasFec As String, dessoc As String, hassoc As String, coope As String, FecFactura As Date, Pb1 As ProgressBar, TipoGasoB As Byte) As Long
'Tipo 0=facturacion normal de la cooperativa correspondiente
'     1=factura de gasoleo bonificado
Dim SQL As String
Dim Rs As ADODB.Recordset

Dim Impuesto As Currency
Dim impbase As Currency
Dim ActSocio As Long
Dim ActForpa As Integer
Dim ActTarje As String 'Long
Dim AntAlbaran As String
Dim AntTarje As String 'Long
Dim AntSocio As Long
Dim AntForpa As Integer
Dim AntTurno As Integer
Dim HayReg As Boolean
Dim v_linea As Integer
Dim FamArtDto As String
Dim IvaArtDto As String
Dim ImporDto As Currency
Dim vCont As CContador
Dim DtoLitro As Currency
Dim CantCombustible As Currency
Dim Codigo As String
Dim baseimpo As Dictionary

Dim CodTipoMov As String

Dim NumError As Long


    On Error GoTo eFacturacion

    NumError = BorramosTemporal(db)


    FamArtDto = "codfamia"
    IvaArtDto = DevuelveDesdeBD("codigiva", "sartic", "codartic", vParamAplic.ArticDto, "N", FamArtDto)
    
    SQL = "select scaalb.codclave, scaalb.codsocio, scaalb.codartic, scaalb.cantidad, scaalb.numlinea,"
    SQL = SQL & " scaalb.preciove, scaalb.importel, scaalb.numalbar, scaalb.fecalbar,"
    SQL = SQL & " scaalb.horalbar, scaalb.codturno, scaalb.codforpa, scaalb.numtarje, scaalb.matricul, "
    SQL = SQL & " scaalb.kilometros, "
    '[Monica]24/08/2015: a�adimos el dto alvic para el regaixo
    SQL = SQL & " scaalb.dtoalvic, "
    '[Monica]28/12/2015: a�adimos el importe vale para el regaixo
    SQL = SQL & " scaalb.importevale "
    
    
    SQL = SQL & " from (scaalb inner join ssocio on scaalb.codsocio = ssocio.codsocio) "
' condicion que tenemos en el datosok
'    sql = sql & " and scoope.tipfactu = 2 )" 'obligatoriamente la cooperativa tiene que tener facturacion ajena
    SQL = SQL & " where scaalb.numfactu = 0 "
    SQL = SQL & " and ssocio.codcoope = " & DBSet(coope, "N")
    If DesFec <> "" Then SQL = SQL & " and scaalb.fecalbar >= '" & Format(CDate(DesFec), FormatoFecha) & "' "
    If HasFec <> "" Then SQL = SQL & " and scaalb.fecalbar <= '" & Format(CDate(HasFec), FormatoFecha) & "' "
    If dessoc <> "" Then SQL = SQL & " and scaalb.codsocio >= " & DBSet(dessoc, "N")
    If hassoc <> "" Then SQL = SQL & " and scaalb.codsocio <= " & DBSet(hassoc, "N")
    
    '[Monica]07/03/2018: seleccionamos los albaranes de articulos que se facturen: facturar = 1
    SQL = SQL & " and scaalb.codartic in (select codartic from sartic where facturar = 1) "

    '[Monica]19/06/2013: si son facturas normales o de gasoleo b
    Select Case TipoGasoB
        Case 0
            SQL = SQL & " and not scaalb.codartic in (select codartic from sartic where tipogaso = 3 union " & _
                                                     "select if(artdto is null, -1, artdto) from sartic where tipogaso = 3  )"
        Case 1
            SQL = SQL & " and scaalb.codartic in (select codartic from sartic where tipogaso = 3 And esdomiciliado = 0 union " & _
                                                 "select if(artdto is null, -1, artdto) from sartic where tipogaso = 3 And esdomiciliado = 0 )"
        Case 2
            SQL = SQL & " and scaalb.codartic in (select codartic from sartic where tipogaso = 3 And esdomiciliado = 1 union " & _
                                                 "select if(artdto is null, -1, artdto) from sartic where tipogaso = 3 And esdomiciliado = 1 )"
    End Select
    
    
    SQL = SQL & " order by scaalb.codsocio, scaalb.codforpa, scaalb.fecalbar, scaalb.horalbar "
    
    Set Rs = db.cursor(SQL)
    HayReg = False
    v_linea = 0
    NumError = 0
    If Not Rs.EOF Then
        Rs.MoveFirst
        AntSocio = Rs!codsocio
        AntAlbaran = Rs!numalbar
        AntForpa = Rs!Codforpa
        AntTurno = Rs!codTurno
        
        Set baseimpo = New Dictionary
        ' cogemos el numero de factura de parametros
        
        Select Case TipoGasoB
            Case 0
                CodTipoMov = Format(CInt(coope), "000")
            Case 1
                CodTipoMov = "C" & Format(CInt(coope), "00")
            Case 2
            
        End Select
        
        numser = ""
        numser = DevuelveDesdeBD("letraser", "stipom", "codtipom", CodTipoMov, "T")
        
        Set vCont = New CContador
        If Not vCont.ConseguirContador(CodTipoMov, True, db) Then Exit Function
        
        
        TotalImp = 0
        
        While Not Rs.EOF And NumError = 0
            HayReg = True
            ActForpa = Rs!Codforpa
            ActSocio = Rs!codsocio
            If (ActForpa <> AntForpa Or ActSocio <> AntSocio) Then
               '  ### [Monica] 05/12/2006
               ' modificacion: si la forma de pago no admite bonificacion no hacemos
               If AdmiteBonificacion(AntForpa) Then
 
                   ' miramos el descuento/litro de socio sobre la cantidad
                   SQL = ""
                   SQL = DevuelveDesdeBD("dtolitro", "ssocio", "codsocio", CStr(AntSocio), "N")
                   DtoLitro = 0
                   If SQL <> "" Then DtoLitro = CCur(SQL)
    
                   If DtoLitro <> 0 Then
                        DtoLitro = DtoLitro * (-1)
                        ImporDto = Round2(CantCombustible * DtoLitro, 2)
                        baseimpo(Val(IvaArtDto)) = DBLet(baseimpo(Val(IvaArtDto)), "N") + DBLet(ImporDto, "N")
                        v_linea = v_linea + 1
                        If NumError = 0 Then NumError = InsertaLineaDescuento(db, numser, vCont.Contador, FecFactura, v_linea, CantCombustible, ImporDto, AntTurno, DtoLitro, AntTarje, 1)
                        If NumError = 0 Then NumError = InsertaLineaDescuentoTemporal(db, CantCombustible, ImporDto)
                   End If
               End If
               
               v_linea = 0
               
               If NumError = 0 Then NumError = InsertCabe(db, baseimpo, vCont.Contador, FecFactura, AntSocio, AntForpa, 1)
               
               Set baseimpo = Nothing
               Set baseimpo = New Dictionary
               TotalImp = 0
               AntForpa = ActForpa
               AntSocio = ActSocio
               AntTurno = Rs!codTurno
               AntTarje = ActTarje
               
               CantCombustible = 0
               
               If Not vCont.ConseguirContador(CodTipoMov, True, db) Then Exit Function
            End If
            
            '-------
            ' tenemos que calcular el impuesto multiplicando cantidad de linea por impuesto por articulo
            Codigo = "codigiva"
            Sql1 = ""
            Sql1 = DevuelveDesdeBD("impuesto", "sartic", "codartic", DBLet(Rs!codartic), "N", Codigo)
            If Sql1 = "" Then
                Impuesto = 0
            Else
                Impuesto = CCur(Sql1) ' Comprueba si es nulo y lo pone a 0 o ""
            End If
            
            If EsArticuloCombustible(Rs!codartic) Then
                TotalImp = TotalImp + Round2((Rs!cantidad * Impuesto), 2)
                CantCombustible = CantCombustible + DBLet(Rs!cantidad, "N")
            End If
            
            baseimpo(Val(Codigo)) = DBLet(baseimpo(Val(Codigo)), "N") + DBLet(Rs!importel, "N")
            v_linea = v_linea + 1
            
            IncrementarProgres Pb1, 1
            
            If NumError = 0 Then NumError = InsertaLineaFactura(db, Rs, numser, vCont.Contador, FecFactura, v_linea, 1)
            If NumError = 0 Then NumError = InsertaLineaFacturaTemporal(db, CStr(Rs!codartic), CStr(Rs!cantidad), CStr(Rs!importel))
            If NumError = 0 Then NumError = BorrarLineaAlbaran(db, Rs!Codclave, True)

            'Siguiente
    '        antfactura = Rs!numfactu
            Rs.MoveNext
        Wend
        If HayReg And NumError = 0 Then
               ' miramos el descuento/litro de socio sobre la cantidad
               
               If AdmiteBonificacion(AntForpa) Then
                    SQL = ""
                    SQL = DevuelveDesdeBD("dtolitro", "ssocio", "codsocio", CStr(AntSocio), "N")
                    DtoLitro = 0
                    If SQL <> "" Then DtoLitro = CCur(SQL)
                    If DtoLitro <> 0 Then
                         DtoLitro = DtoLitro * (-1)
                         ImporDto = Round2(CantCombustible * DtoLitro, 2)
                         baseimpo(Val(IvaArtDto)) = DBLet(baseimpo(Val(IvaArtDto)), "N") + DBLet(ImporDto, "N")
                         v_linea = v_linea + 1
                         If NumError = 0 Then NumError = InsertaLineaDescuento(db, numser, vCont.Contador, FecFactura, v_linea, CantCombustible, ImporDto, AntTurno, DtoLitro, AntTarje, 1)
                        If NumError = 0 Then NumError = InsertaLineaDescuentoTemporal(db, CantCombustible, ImporDto)
                    End If
               End If
               If NumError = 0 Then NumError = InsertCabe(db, baseimpo, vCont.Contador, FecFactura, AntSocio, AntForpa, 1)
               
            '[Monica]07/03/2018: eliminamos los albaranes de los articulos a no facturar
            If NumError = 0 Then
                NumError = BorrarAlbaranesArtNoFacturan(db, DesFec, HasFec, dessoc, hassoc, coope, TipoGasoB)
            End If
        End If
        ' hemos de incluir la factura global de la cooperativa, partiendo de la temporal que tenemos grabada.
        If NumError = 0 Then
            NumError = InsertarFacturaGlobal(db, coope, FecFactura, 0, Pb1, TipoGasoB)
        End If
        
 
        
    End If
    
eFacturacion:
    FacturacionAjena = NumError
    Exit Function
End Function

Private Function BorrarAlbaranesArtNoFacturan(db As BaseDatos, DesFec As String, HasFec As String, dessoc As String, hassoc As String, coope As String, TipoGasoB As Byte) As Long
Dim SQL As String


On Error GoTo eBorrarAlbaranesArtNoFacturan


    SQL = "delete from scaalb "
    SQL = SQL & " where scaalb.numfactu = 0 "
    SQL = SQL & " and scaalb.codsocio in (select codsocio from ssocio where ssocio.codcoope = " & DBSet(coope, "N") & ") "
    If DesFec <> "" Then SQL = SQL & " and scaalb.fecalbar >= '" & Format(CDate(DesFec), FormatoFecha) & "' "
    If HasFec <> "" Then SQL = SQL & " and scaalb.fecalbar <= '" & Format(CDate(HasFec), FormatoFecha) & "' "
    If dessoc <> "" Then SQL = SQL & " and scaalb.codsocio >= " & DBSet(dessoc, "N")
    If hassoc <> "" Then SQL = SQL & " and scaalb.codsocio <= " & DBSet(hassoc, "N")
    
    '[Monica]07/03/2018: seleccionamos los articulos que no se facturan: facturar = 0
    SQL = SQL & " and scaalb.codartic in (select codartic from sartic where facturar = 0) "
    
    
    '[Monica]19/06/2013: si son facturas normales o de gasoleo b
    Select Case TipoGasoB
        Case 0
            SQL = SQL & " and not scaalb.codartic in (select codartic from sartic where tipogaso = 3 union " & _
                                                     "select if(artdto is null, -1, artdto) from sartic where tipogaso = 3 )"
        Case 1
            SQL = SQL & " and scaalb.codartic in (select codartic from sartic where tipogaso = 3 And esdomiciliado = 0 union " & _
                                                 "select if(artdto is null, -1, artdto) from sartic where tipogaso = 3 And esdomiciliado = 0 )"
        Case 2
            SQL = SQL & " and scaalb.codartic in (select codartic from sartic where tipogaso = 3 And esdomiciliado = 1 union " & _
                                                 "select if(artdto is null, -1, artdto) from sartic where tipogaso = 3 And esdomiciliado = 1 )"
    End Select
    
    NumError = db.ejecutar(SQL)

eBorrarAlbaranesArtNoFacturan:
    If Err.Number <> 0 Then NumError = Err.Number
    
    BorrarAlbaranesArtNoFacturan = NumError
End Function


' funcion que nos permite insertar la factura global que se le hace a la cooperativa
' se utiliza en la facturacion ajena del Regaixo
Private Function InsertarFacturaGlobal(db As BaseDatos, coope As String, FecFactura As Date, Tipo As Byte, Optional Pb1 As ProgressBar, Optional TipoGasoB As Byte) As Long
' tipo = 0 factura de gasoleo a la cooperativa
' tipo = 1 factura de bonificacion a la cooperativa

'TipoGasoB = 0 factura de gasoleo a la cooperativa (solo si el tipo = 0  factura de gasoleo)
'TipoGasoB = 1 factura de gasoleo domiciliado a la cooperativa (solo si el tipo = 0  factura de gasoleo)

Dim SQL As String
Dim vCont As CContador
Dim Sql1 As String
Dim Rs As ADODB.Recordset
Dim socio As Long
Dim Numtarje As String

Dim I As Integer
Dim Imptot(2)
Dim Tipiva(2)
Dim Impbas(2)
Dim ImpIva(2)
Dim PorIva(2)
Dim TotFac
Dim NumCoop As String
Dim baseimpo As Dictionary
Dim Forpa As String
Dim Codigo As String
Dim preciove As Currency
Dim Serie As String
Dim articulo As String

'importe1 = articulo
'importe2 = cantidad
'importe3 = importel


On Error GoTo eInsertarFacturaGlobal

    '[Monica]02/01/2017: a�adida la condicion de que el importe2 del articulo sea <> 0 para que no inserte la linea con importe 0
    SQL = "select importe1, importe2, importe3 from tmpinformes where codusu = " & vSesion.Codigo & " and importe2 <> 0 order by 1"
    Set Rs = db.cursor(SQL)
    
    ' INSERTAMOS LAS LINEAS DE LA FACTURA, UNA LINEA POR CADA ARTICULO
    Set baseimpo = Nothing
    Set baseimpo = New Dictionary
    TotalImp = 0
    CantCombustible = 0
    v_linea = 0
    
    Set vCont = New CContador
    Serie = ""
    If Tipo = 0 Then
        '[Monica]19/06/2013: dependiendo del tipo de gasoleo que sea domiciliado o no
        If TipoGasoB = 0 Then
            If Not vCont.ConseguirContador("FAG", True, db) Then
                InsertarFacturaGlobal = -1
                Exit Function
            End If
            Serie = DevuelveDesdeBDNew(cPTours, "stipom", "letraser", "codtipom", "FAG", "T")
        Else
            If Not vCont.ConseguirContador("FGB", True, db) Then
                InsertarFacturaGlobal = -1
                Exit Function
            End If
            Serie = DevuelveDesdeBDNew(cPTours, "stipom", "letraser", "codtipom", "FGB", "T")
        End If
    Else
        If Not vCont.ConseguirContador("FAB", True, db) Then
            InsertarFacturaGlobal = -1
            Exit Function
        End If
        Serie = DevuelveDesdeBDNew(cPTours, "stipom", "letraser", "codtipom", "FAB", "T")
    End If
    
    ' dependiendo de la cooperativa se asignar� la factura a un socio u otro
    ' esto lo parametrizaremos si hay otra cooperativa que funciona igual
    If coope = 1 Then socio = 3007
    If coope = 2 Then socio = 3008
    
    Numtarje = ""
    Numtarje = DevuelveDesdeBDNew(cPTours, "starje", "numtarje", "codsocio", CStr(socio), "N")
    Codforpa = ""
    Codforpa = DevuelveDesdeBDNew(cPTours, "ssocio", "codforpa", "codsocio", CStr(socio), "N")
    
    While Not Rs.EOF And NumError = 0
        '-------
        ' tenemos que calcular el impuesto multiplicando cantidad de linea por impuesto por articulo
        If Tipo = 1 Then
            '11/09/08: antes: articulo = "00" & mid(rs!importe1, 3, 4)
            articulo = Rs!Importe1 ' "00" & Mid(Rs!Importe1, 3, 4)
        Else
            articulo = Rs!Importe1
        End If
        Codigo = "codigiva"
        Sql1 = ""
        Sql1 = DevuelveDesdeBD("impuesto", "sartic", "codartic", DBLet(articulo), "N", Codigo)
        If Sql1 = "" Then
            Impuesto = 0
        Else
            Impuesto = CCur(Sql1) ' Comprueba si es nulo y lo pone a 0 o ""
        End If
        
        If EsArticuloCombustible(DBLet(Rs!Importe1)) Then
            TotalImp = TotalImp + Round2((DBLet(Rs!Importe2) * Impuesto), 2)
            CantCombustible = CantCombustible + DBLet(Rs!Importe2, "N")
        End If
        
        baseimpo(Val(Codigo)) = DBLet(baseimpo(Val(Codigo)), "N") + DBLet(Rs!importe3, "N")
        v_linea = v_linea + 1
        
        IncrementarProgres Pb1, 1
        preciove = 0
        If DBLet(Rs!Importe2) <> 0 Then
            preciove = Round2(DBLet(Rs!importe3) / DBLet(Rs!Importe2), 3)
        End If
        ' insertamos la linea de factura
        SQL = "INSERT into slhfac (letraser, numfactu, fecfactu, numlinea, numalbar, " & _
                "fecalbar, horalbar, codturno, numtarje, codartic, " & _
                "cantidad, preciove, implinea) " & _
                "values " & _
                "(" & db.Texto(Serie) & "," & db.numero(vCont.Contador) & "," & db.Fecha(FecFactura) & "," & db.numero(v_linea) & "," & db.Texto("COOP.") & "," & _
                db.Fecha(FecFactura) & "," & db.fechahora(FecFactura & " " & Format(Time, "hh:mm:ss")) & ",1," & db.numero(Numtarje) & "," & db.numero(Rs!Importe1) & "," & _
                db.numero(Rs!Importe2) & "," & db.numero(preciove) & "," & db.numero(Rs!importe3) & ")"
           
        NumError = db.ejecutar(SQL)

        Rs.MoveNext
     
     Wend
     

    ' finalmente insertamos la cabecera de factura
    ' inicializamos los importes de los totales de la cabecera
    TotFac = 0
    For I = 0 To 2
         Tipiva(I) = Null
         Imptot(I) = Null
         Impbas(I) = Null
         ImpIva(I) = Null
         PorIva(I) = Null
    Next I
    
    For I = 0 To baseimpo.Count - 1
        If I <= 2 Then
            Tipiva(I) = baseimpo.Keys(I)
            Imptot(I) = baseimpo.Items(I)
            PorIva(I) = DevuelveDesdeBDNew(cConta, "tiposiva", "porceiva", "codigiva", CStr(Tipiva(I)), "N")
            Impbas(I) = Round2(Imptot(I) / (1 + (PorIva(I) / 100)), 2)
            ImpIva(I) = Imptot(I) - Impbas(I)
            TotFac = TotFac + Imptot(I)
        End If
    Next I
    
    NumCoop = coope
    
    SQL = "INSERT into schfac "
    SQL = SQL & "(letraser, numfactu, fecfactu, codsocio, codcoope, " & _
           "codforpa, baseimp1, baseimp2, baseimp3, impoiva1, " & _
           "impoiva2, impoiva3, tipoiva1, tipoiva2, tipoiva3, " & _
           "porciva1, porciva2, porciva3, totalfac, impuesto, " & _
           "intconta)" & _
           "values " & _
           "(" & db.Texto(Serie) & "," & db.numero(vCont.Contador) & "," & db.Fecha(FecFactura) & "," & db.numero(socio) & "," & db.numero(NumCoop) & "," & _
           db.numero(Forpa) & "," & db.numero(Impbas(0)) & "," & db.numero(Impbas(1)) & "," & db.numero(Impbas(2)) & "," & db.numero(ImpIva(0)) & "," & _
           db.numero(ImpIva(1)) & "," & db.numero(ImpIva(2)) & "," & db.numero(Tipiva(0)) & "," & db.numero(Tipiva(1)) & "," & db.numero(Tipiva(2)) & "," & _
           db.numero(PorIva(0)) & "," & db.numero(PorIva(1)) & "," & db.numero(PorIva(2)) & "," & db.numero(TotFac) & "," & db.numero(TotalImp) & "," & _
           "0" & ")"
    
    NumError = db.ejecutar(SQL)

eInsertarFacturaGlobal:
    If Err.Number <> 0 Then NumError = Err.Number
    
    InsertarFacturaGlobal = NumError

End Function


Public Function BorramosTemporal(ByRef db As BaseDatos) As Long
Dim SQL As String

    SQL = "delete from tmpinformes where codusu = " & vSesion.Codigo
    BorramosTemporal = db.ejecutar(SQL)
    
End Function


Public Function FacturacionAbonoCliente(ByRef db As BaseDatos, codEmpre As Currency, Cooperativa As String, desdef As String, hastaf As String, fecFac As String, Serie As String) As Long
Dim SQL As String
Dim Rs As ADODB.Recordset
Dim Rs2 As ADODB.Recordset
Dim sql2 As String
Dim AntSocio As String, ActSocio As String
Dim vsocio As CSocio
Dim cantidad As Currency
Dim Importe As Currency
Dim Precio As Currency
Dim b As Boolean
Dim Linea As Integer
Dim vCont As CContador
Dim NumError As Long
Dim baseimpo As Dictionary
Dim HayReg As Byte
Dim Sql1 As String
Dim Impuesto As Currency
'Dim TotalImp As Currency
Dim CantCombustible As Currency
Dim v_linea As Integer
Dim Codigo As String
Dim Hora As String
Dim ArtDto As String
Dim FamDto As String

On Error GoTo eFacturacionAbonoCliente

    FacturacionAbonoCliente = 0
    '03/05/07 antes la condicion era: schfac.letraser <> dbset(serie,"T")
    SQL = "select schfac.codsocio, slhfac.codartic, tmpinformes.precio1, sum(cantidad) " & _
          " from schfac, slhfac, sligru, tmpinformes, ssocio " & _
          " where sligru.codempre = " & DBSet(codEmpre, "N") & " and " & _
              "ssocio.codcoope = " & DBSet(Cooperativa, "N") & " and " & _
              "schfac.letraser = " & DBSet(Serie, "T") & " and " & _
              "tmpinformes.codusu = " & vSesion.Codigo
              
    If desdef <> "" Then SQL = SQL & " and slhfac.fecfactu >= " & DBSet(desdef, "F")
    If hastaf <> "" Then SQL = SQL & " and slhfac.fecfactu <= " & DBSet(hastaf, "F")
    
    SQL = SQL & " and sligru.codsocio = schfac.codsocio and schfac.codsocio = ssocio.codsocio  " & _
                " and schfac.letraser = slhfac.letraser and schfac.numfactu = slhfac.numfactu and schfac.fecfactu = slhfac.fecfactu " & _
                " and slhfac.codartic = tmpinformes.codigo1 " & _
                " group by 1, 2, 3 " & _
                " order by 1, 2, 3 "
    
    Set Rs = db.cursor(SQL)
    
    If Rs.EOF Then Exit Function
    
    AntSocio = Rs.Fields(0).Value
    ActSocio = Rs.Fields(0).Value
    Linea = 0
    Set vCont = New CContador
    
    If Not vCont.ConseguirContador("FAB", True, db) Then Exit Function
    
    numser = ""
    numser = DevuelveDesdeBD("letraser", "stipom", "codtipom", "FAB", "T")
    
    
    Hora = Format(Now, "hh:mm:ss")
    
    
    Set baseimpo = New Dictionary
    
    ' obtenemos la familia de descuento para saber sacar el articulo de dto
    FamDto = ""
    FamDto = DevuelveDesdeBDNew(cPTours, "sfamia", "codfamia", "tipfamia", 2, "N")
    
    TotalImp = 0
    
    HayReg = 0
    While Not Rs.EOF And NumError = 0
        HayReg = 1
        ActSocio = Rs.Fields(0).Value
        If ActSocio <> AntSocio Then
            Set vsocio = New CSocio
            If vsocio.LeerDatos(AntSocio) Then
                 NumError = InsertCabe(db, baseimpo, vCont.Contador, CDate(fecFac), vsocio.Codigo, vsocio.ForPago, 0)
                 AntSocio = ActSocio
                 TotalImp = 0
            End If
            Set baseimpo = Nothing
            Set baseimpo = New Dictionary
        
            If Not vCont.ConseguirContador("FAB", True, db) Then Exit Function
            
            v_linea = 0
        End If

        ArtDto = Format(Rs.Fields(1).Value, "000000")
        ArtDto = Format(FamDto, "00") & Mid(ArtDto, 3, 5)
        
        
        Codigo = "codigiva"
        Sql1 = ""
        Sql1 = DevuelveDesdeBD("impuesto", "sartic", "codartic", DBLet(ArtDto, "N"), "N", Codigo) ' 04/05/07 antes era rs!codartic
        If Sql1 = "" Then
            Impuesto = 0
        Else
            Impuesto = CCur(Sql1) ' Comprueba si es nulo y lo pone a 0 o ""
        End If
        
        If EsArticuloCombustible(Rs!codartic) Then
        ' restamos porque estamos en abono, en la facturacion se suma
            TotalImp = TotalImp + Round2((Rs.Fields(3).Value * Impuesto * (-1)), 2)
            CantCombustible = CantCombustible + DBLet(Rs.Fields(3).Value, "N")
        End If
        
        Precio = Rs.Fields(2).Value * (-1)
        Importe = Round2(Precio * Rs.Fields(3).Value, 2)
        
        baseimpo(Val(Codigo)) = DBLet(baseimpo(Val(Codigo)), "N") + DBLet(Importe, "N")
        v_linea = v_linea + 1
        

        NumError = InsertaLineaFacturaAbono(db, Rs, numser, vCont.Contador, CDate(fecFac), Hora, v_linea, Rs.Fields(3).Value, Precio, Importe, ArtDto, 0)
        Rs.MoveNext
    Wend
    
    ' insertamos la ultima cabecera de factura
    If HayReg = 1 And NumError = 0 Then
        Set vsocio = New CSocio
        If vsocio.LeerDatos(ActSocio) Then
             NumError = InsertCabe(db, baseimpo, vCont.Contador, CDate(fecFac), vsocio.Codigo, vsocio.ForPago, 0)
             AntSocio = ActSocio
        End If
        Set baseimpo = Nothing
    End If

eFacturacionAbonoCliente:
    If Err.Number <> 0 Or NumError <> 0 Then
        FacturacionAbonoCliente = 1
    End If
End Function



Public Function InsertaLineaFacturaAbono(ByRef db As BaseDatos, ByRef Rs As ADODB.Recordset, numser As String, numFac As Long, fecFac As Date, Hora As String, Linea As Integer, cantidad As Currency, Precio As Currency, Importe As Currency, ArtDto As String, Tipo As Byte) As Long
Dim Numtarje As String
' tipo = 0 facturacion
' tipo = 1 facturacion ajena

    Dim SQL As String
    Dim ImpLinea As Currency
    
    On Error GoTo eInsertaLineaFacturaAbono
    MensError = ""
    
    If Tipo = 0 Then
        SQL = "INSERT into slhfac "
    Else
        SQL = "INSERT into slhfacr "
    End If
     
    Numtarje = ""
    Numtarje = DevuelveDesdeBDNew(cPTours, "starje", "numtarje", "codsocio", Rs!codsocio, "N")
     
     SQL = SQL & "(letraser, numfactu, fecfactu, numlinea, numalbar, " & _
           "fecalbar, horalbar, codturno, numtarje, codartic, " & _
           "cantidad, preciove, implinea) " & _
           "values " & _
           "(" & db.Texto(numser) & "," & db.numero(numFac) & "," & db.Fecha(fecFac) & "," & db.numero(Linea) & ",'BONIFICA'," & _
           db.Fecha(fecFac) & "," & db.fechahora(fecFac & " " & Format(Hora, "hh:mm:ss")) & "," & db.numero(1) & "," & db.numero(Numtarje) & "," & db.numero(ArtDto) & "," & _
           db.numero(cantidad) & "," & db.numero(Precio) & "," & db.numero(Importe) & ")"
           
    InsertaLineaFacturaAbono = db.ejecutar(SQL)

eInsertaLineaFacturaAbono:
    If Err.Number <> 0 Or InsertaLineaFacturaAbono <> 0 Then
        MensError = "Error en la inserci�n de la l�nea de factura " & numFac & " en el articulo " & Rs!codartic
        If InsertaLineaFacturaAbono = 0 Then InsertaLineaFacturaAbono = 1
    End If
    
End Function



Public Function FacturacionAbonoSocio(desdesoc As String, hastasoc As String, desdefec As String, hastafec As String, SerBonif As String, fecFac As String, Cooperativa As String, ByRef Pb1 As ProgressBar) As Boolean
Dim SQL As String
Dim sql2 As String
Dim Rs As ADODB.Recordset
Dim ActCodsocio As String
Dim ActCodartic As String
Dim AntCodsocio As String
Dim AntCodartic As String
Dim HayReg As Byte
Dim v_linea As Integer
Dim NumError As Long
Dim BONIFICA As Currency
Dim b As Boolean
Dim db As BaseDatos
Dim NRegs As Integer
Dim Codigo As String
Dim Hora As String
Dim vsocio As CSocio
Dim ArtDto As String


     On Error GoTo eFacturacionAbonoSocio


     Set db = New BaseDatos
     db.abrir vSesion.CadenaConexion, "root", "aritel"
     db.Tipo = "MYSQL"
     db.AbrirTrans

    NumError = 0
    
    NumError = BorramosTemporal(db)
    
    ' realizamos la facturacion
    SQL = "select schfacr.codsocio, slhfacr.codartic, sum(cantidad) "
    SQL = SQL & "from schfacr, slhfacr, ssocio, sartic, sfamia "
    SQL = SQL & " where sfamia.tipfamia = 1 " ' unicamente carburantes
    SQL = SQL & " and sartic.bonigral <> 0 "
    SQL = SQL & " and schfacr.letraser <> " & DBSet(SerBonif, "T")
    If desdesoc <> "" Then SQL = SQL & " and schfacr.codsocio >= " & DBSet(desdesoc, "N")
    If hastasoc <> "" Then SQL = SQL & " and schfacr.codsocio <= " & DBSet(hastasoc, "N")
    If desdefec <> "" Then SQL = SQL & " and slhfacr.fecfactu >= " & DBSet(desdefec, "F")
    If hastafec <> "" Then SQL = SQL & " and slhfacr.fecfactu <= " & DBSet(hastafec, "F")
    SQL = SQL & " and ssocio.codcoope = " & DBSet(Cooperativa, "N")
    SQL = SQL & " and schfacr.codsocio = ssocio.codsocio "
    SQL = SQL & " and sfamia.codfamia = sartic.codfamia "
    SQL = SQL & " and slhfacr.codartic = sartic.codartic "
    SQL = SQL & " and slhfacr.letraser = schfacr.letraser and slhfacr.numfactu = schfacr.numfactu and schfacr.fecfactu = slhfacr.fecfactu "
    SQL = SQL & " GROUP BY schfacr.codsocio, slhfacr.codartic "
    SQL = SQL & " ORDER BY schfacr.codsocio, slhfacr.codartic "

    Set Rs = db.cursor(SQL)
    HayReg = False
    Set Rs = db.cursor(SQL)
    
    If Rs.EOF Then
        FacturacionAbonoSocio = True
        Exit Function
    End If
    
    AntCodsocio = Rs.Fields(0).Value
    ActCodsocio = Rs.Fields(0).Value
    Linea = 0
    
    Set vCont = New CContador
    If Not vCont.ConseguirContador("B" & Format(Cooperativa, "00"), True, db) Then Exit Function
    
    numser = ""
    numser = DevuelveDesdeBD("letraser", "stipom", "codtipom", "B" & Format(Cooperativa, "00"), "T")
    
    
    Hora = Format(Now, "hh:mm:ss")
    
    
    Set baseimpo = New Dictionary
    
    ' obtenemos la familia de descuento para saber sacar el articulo de dto
    FamDto = ""
    FamDto = DevuelveDesdeBDNew(cPTours, "sfamia", "codfamia", "tipfamia", 2, "N")
        
    TotalImp = 0
    
    HayReg = 0
    While Not Rs.EOF And NumError = 0
        HayReg = 1
        IncrementarProgres Pb1, 1
        ActCodsocio = Rs.Fields(0).Value
        If ActCodsocio <> AntCodsocio Then
            Set vsocio = New CSocio
            If vsocio.LeerDatos(AntCodsocio) Then
                 NumError = InsertCabe(db, baseimpo, vCont.Contador, CDate(fecFac), vsocio.Codigo, vsocio.ForPago, 1)
                 AntCodsocio = ActCodsocio
            End If
            Set baseimpo = Nothing
            Set baseimpo = New Dictionary
        
            If Not vCont.ConseguirContador("B" & Format(Cooperativa, "00"), True, db) Then Exit Function
        
        End If

        Codigo = "codigiva"
        Sql1 = ""
        Sql1 = DevuelveDesdeBD("impuesto", "sartic", "codartic", DBLet(Rs!codartic), "N", Codigo)
        If Sql1 = "" Then
            Impuesto = 0
        Else
            Impuesto = CCur(Sql1) ' Comprueba si es nulo y lo pone a 0 o ""
        End If
        
        ArtDto = Format(Rs.Fields(1).Value, "000000")
        ArtDto = Format(FamDto, "00") & Mid(ArtDto, 3, 5)
        
        If EsArticuloCombustible(ArtDto) Then ' antes rs!codartic
        ' restamos porque estamos en abono, en la facturacion se suma
            TotalImp = TotalImp + Round2((Rs.Fields(2).Value * Impuesto * (-1)), 2)
            CantCombustible = CantCombustible + DBLet(Rs.Fields(2).Value, "N")
        End If
        
        Precio = ""
        Precio = DevuelveDesdeBDNew(cPTours, "sartic", "bonigral", "codartic", Rs.Fields(1).Value, "N")
        
        v_precio = CCur(Precio) * (-1)
        Importe = Round2(v_precio * Rs.Fields(2).Value, 2)
        
        ' insertamos en la temporal para hacer la factura a la cooperativa
        If NumError = 0 Then ' a�adida condicion 12/07/2007
            NumError = InsertaLineaFacturaTemporal(db, CStr(ArtDto), CStr(Rs.Fields(2).Value), CStr(Importe))
        End If
        
        baseimpo(Val(Codigo)) = DBLet(baseimpo(Val(Codigo)), "N") + DBLet(Importe, "N")
        v_linea = v_linea + 1
        
        If NumError = 0 Then ' a�adida condicion 12/07/2007
            NumError = InsertaLineaFacturaAbono(db, Rs, numser, vCont.Contador, CDate(fecFac), Hora, v_linea, Rs.Fields(2).Value, CCur(v_precio), CCur(Importe), CStr(ArtDto), 1)
        End If
        Rs.MoveNext
    Wend
    
    ' insertamos la ultima cabecera de factura
    If HayReg = 1 And NumError = 0 Then ' a�adida condicion 12/07/2007 and numerror = 0
        Set vsocio = New CSocio
        If vsocio.LeerDatos(ActCodsocio) Then
             NumError = InsertCabe(db, baseimpo, vCont.Contador, CDate(fecFac), vsocio.Codigo, vsocio.ForPago, 1)
             AntCodsocio = ActCodsocio
             NumError = InsertarFacturaGlobal(db, Cooperativa, CDate(fecFac), 1)
        End If
        Set baseimpo = Nothing
    End If

eFacturacionAbonoSocio:
    If Err.Number <> 0 Or NumError Then
        If Err.Number <> 0 Then
            FacturacionAbonoSocio = Err.Number
        Else
            FacturacionAbonoSocio = NumError
        End If
        db.RollbackTrans
    Else
        FacturacionAbonoSocio = 0
        db.CommitTrans
    End If
    Set db = Nothing
End Function

Public Function CrearFacturaRectificativa(letraser As String, numfactu As String, fecfactu As String, observac As String, NuevoCliente As String, NuevaFecFactu As String, RecuperaAlbaranes As Boolean) As Boolean
Dim SQL As String
Dim sql2 As String
Dim Rs As ADODB.Recordset
Dim Rs2 As ADODB.Recordset
Dim db As BaseDatos
Dim vCont As CContador
Dim vsocio As CSocio
Dim Contabilizada As Byte
Dim NumError As Long
Dim caderr As String
Dim numser As String

Dim Traba As String
Dim Codclave As Long

     On Error GoTo eCrearFacturaRectificativa

     Set db = New BaseDatos
     
     db.abrir vSesion.CadenaConexion, "root", "aritel"
     db.Tipo = "MYSQL"
     db.Con = Conn
     db.AbrirTrans
     ConnConta.BeginTrans
     NumError = 0

     SQL = "select * from schfac where letraser = " & DBSet(letraser, "T") & " and numfactu = " & DBSet(numfactu, "N")
     SQL = SQL & " and fecfactu = " & DBSet(fecfactu, "F")
     
     Set Rs = db.cursor(SQL)
    
       
     ' factura en negativo
     Set vCont = New CContador
     If Not vCont.ConseguirContador("FAR", True, db) Then
        CrearFacturaRectificativa = -1
        db.RollbackTrans
        ConnConta.RollbackTrans
        Exit Function
     End If
     
     numser = ""
     numser = DevuelveDesdeBD("letraser", "stipom", "codtipom", "FAR", "T")
    
     If numser = "" Then
        MsgBox "La Letra de Serie de la Factura Rectificativa tiene que tener un valor." & vbCrLf & vbCrLf & "Revise.", vbExclamation
        CrearFacturaRectificativa = -1
        db.RollbackTrans
        ConnConta.RollbackTrans
        Exit Function
     End If
    
    
     Contabilizada = DBLet(Rs!intconta, "N")
     Set vsocio = New CSocio
     
     If Not vsocio.LeerDatos(Rs!codsocio) Then
        NumError = -1
     Else
         sql2 = "insert into schfac (letraser, numfactu, fecfactu, codsocio, codcoope, codforpa, "
         sql2 = sql2 & "baseimp1, baseimp2, baseimp3, impoiva1, impoiva2, impoiva3, tipoiva1,"
         sql2 = sql2 & "tipoiva2, tipoiva3, porciva1, porciva2, porciva3, totalfac, impuesto,"
         sql2 = sql2 & "intconta, observac, rectif_letraser, rectif_numfactu, rectif_fecfactu) values ("
         sql2 = sql2 & DBSet(numser, "T") & "," & DBSet(vCont.Contador, "N") & ","
         sql2 = sql2 & DBSet(NuevaFecFactu, "F") & "," & DBSet(Rs!codsocio, "N") & "," & DBSet(Rs!codcoope, "N") & ","
         sql2 = sql2 & DBSet(Rs!Codforpa, "N") & ","
         
'[Monica]16/10/2013: no puede ser nulo
'         If DBLet(Rs!baseimp1, "N") <> 0 Then
            sql2 = sql2 & DBSet(DBLet(Rs!baseimp1, "N") * (-1), "N") & ","
'         Else
'            Sql2 = Sql2 & "null,"
'         End If
         If DBLet(Rs!baseimp2, "N") <> 0 Then
             sql2 = sql2 & DBSet(DBLet(Rs!baseimp2, "N") * (-1), "N") & ","
         Else
            sql2 = sql2 & "null,"
         End If
         If DBLet(Rs!baseimp3, "N") <> 0 Then
             sql2 = sql2 & DBSet(DBLet(Rs!baseimp3, "N") * (-1), "N") & ","
         Else
            sql2 = sql2 & "null,"
         End If
'[Monica]16/10/2013: no puede ser nulo
'         If DBLet(Rs!impoiva1, "N") <> 0 Then
             sql2 = sql2 & DBSet(DBLet(Rs!impoiva1, "N") * (-1), "N") & ","
'         Else
'            Sql2 = Sql2 & "null,"
'         End If
         If DBLet(Rs!impoiva2, "N") <> 0 Then
             sql2 = sql2 & DBSet(DBLet(Rs!impoiva2, "N") * (-1), "N") & ","
         Else
            sql2 = sql2 & "null,"
         End If
         If DBLet(Rs!impoiva3, "N") <> 0 Then
             sql2 = sql2 & DBSet(DBLet(Rs!impoiva3, "N") * (-1), "N") & ","
         Else
            sql2 = sql2 & "null,"
         End If
         sql2 = sql2 & DBSet(Rs!TipoIVA1, "N") & ","
         sql2 = sql2 & DBSet(Rs!TipoIVA2, "N") & ","
         sql2 = sql2 & DBSet(Rs!TipoIVA3, "N") & ","
         sql2 = sql2 & DBSet(Rs!porciva1, "N") & ","
         sql2 = sql2 & DBSet(Rs!porciva2, "N") & ","
         sql2 = sql2 & DBSet(Rs!porciva3, "N") & ","
         sql2 = sql2 & DBSet(DBLet(Rs!TotalFac, "N") * (-1), "N") & ","
         sql2 = sql2 & DBSet(DBLet(Rs!Impuesto, "N") * (-1), "N") & ","
         sql2 = sql2 & DBSet(0, "N") & ","
         sql2 = sql2 & DBSet(observac, "T") & ","
'         sql2 = sql2 & DBSet(RS!intconta, "N") & ","
         sql2 = sql2 & DBSet(letraser, "T") & ","
         sql2 = sql2 & DBSet(numfactu, "N") & ","
         sql2 = sql2 & DBSet(fecfactu, "F") & ")"
         
         caderr = "Insertando cabecera de factura rectificativa:"
         NumError = db.ejecutar2(sql2, caderr)
         
         Rs.Close
         
         SQL = "select * from slhfac where letraser = " & DBSet(letraser, "T") & " and numfactu = " & DBSet(numfactu, "N")
         SQL = SQL & " and fecfactu = " & DBSet(fecfactu, "F")
         
'         Set RS = New adodb.Recordset
'         RS.Open sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
         Set Rs = db.cursor(SQL)
         
         caderr = "Insertando lineas de factura rectificativa:"
         
         While Not Rs.EOF And (NumError = 0)
            sql2 = "insert into slhfac (letraser,numfactu,fecfactu,numlinea,numalbar,fecalbar,horalbar,"
            sql2 = sql2 & "codturno,numtarje,codartic,cantidad,preciove,implinea) values ("
            sql2 = sql2 & DBSet(numser, "T") & "," & DBSet(vCont.Contador, "N") & "," & DBSet(NuevaFecFactu, "F") & ","
            sql2 = sql2 & DBSet(Rs!NumLinea, "N") & "," & DBSet(Rs!numalbar, "T") & "," & DBSet(Rs!fecAlbar, "F") & ","
            sql2 = sql2 & DBSet(Rs!horalbar, "FH") & "," & DBSet(Rs!codTurno, "N") & "," & DBSet(Rs!Numtarje, "N") & ","
            sql2 = sql2 & DBSet(Rs!codartic, "N") & "," & DBSet(DBLet(Rs!cantidad, "N") * (-1), "N") & ","
            sql2 = sql2 & DBSet(Rs!preciove, "N") & "," & DBSet(DBLet(Rs!ImpLinea, "N") * (-1), "N") & ")"
            
            NumError = db.ejecutar2(sql2, caderr)
            
            Rs.MoveNext
         Wend
    End If
    
    If NumError = 0 Then
        '[Monica]18/01/2013: recuperamos los albaranes de la factura
        If RecuperaAlbaranes Then
             SQL = "select schfac.codsocio, schfac.codforpa, slhfac.* from schfac inner join slhfac on schfac.letraser = slhfac.letraser and schfac.numfactu = slhfac.numfactu and schfac.fecfactu = slhfac.fecfactu "
             SQL = SQL & " where schfac.letraser = " & DBSet(letraser, "T") & " and schfac.numfactu = " & DBSet(numfactu, "N")
             SQL = SQL & " and schfac.fecfactu = " & DBSet(fecfactu, "F")
             SQL = SQL & " order by numlinea "
             
             Set Rs2 = db.cursor(SQL)
             
             caderr = "Insertando lineas de factura en albaranes:"
             
             Traba = DevuelveValor("select min(codtraba) from straba")
             
             While Not Rs2.EOF And (NumError = 0)
                ' insertamos en la tabla de albaranes
                Codclave = SugerirCodigoSiguienteStr("scaalb", "codclave")
                 
                If Rs2!numalbar <> "BONIFICA" Then
                    sql2 = "insert into scaalb (codclave,codsocio,numtarje,numalbar,fecalbar,horalbar,codturno,codartic,cantidad,preciove,"
                    sql2 = sql2 & "importel,codforpa,matricul,codtraba,numfactu,numlinea,declaradogp,precioinicial) values ("
                    sql2 = sql2 & DBSet(Codclave, "N") & "," & DBSet(Rs2!codsocio, "N") & "," & DBSet(Rs2!Numtarje, "N") & ","
                    sql2 = sql2 & DBSet(Rs2!numalbar, "N") & "," & DBSet(Rs2!fecAlbar, "F") & ","
                    sql2 = sql2 & DBSet(Rs2!horalbar, "FH") & "," & DBSet(Rs2!codTurno, "N") & ","
                    sql2 = sql2 & DBSet(Rs2!codartic, "N") & "," & DBSet(Rs2!cantidad, "N") & ","
                    sql2 = sql2 & DBSet(Rs2!preciove, "N") & "," & DBSet(Rs2!ImpLinea, "N") & ","
                    sql2 = sql2 & DBSet(Rs2!Codforpa, "N") & "," & DBSet(Rs2!matricul, "T") & ","
                    sql2 = sql2 & DBSet(Traba, "N") & ",0,0,0," & DBSet(Rs2!precioinicial, "N") & ")"
                    
                    NumError = db.ejecutar2(sql2, caderr)
                    
                End If
                
                Rs2.MoveNext
             Wend
        
            Set Rs2 = Nothing
            
        Else
             'factura para el nuevo cliente si lo hay
            If NuevoCliente <> "" Then ' and b
                 SQL = "select * from schfac where letraser = " & DBSet(letraser, "T") & " and numfactu = " & DBSet(numfactu, "N")
                 SQL = SQL & " and fecfactu = " & DBSet(fecfactu, "F")
                 
                 Set Rs = db.cursor(SQL)
                  
                 
                 Set vCont = New CContador
                 If Not vCont.ConseguirContador("FAG", True, db) Then Exit Function
                 
                 numser = ""
                 numser = DevuelveDesdeBD("letraser", "stipom", "codtipom", "FAG", "T")
                
                 Contabilizada = DBLet(Rs!intconta, "N")
                 Set vsocio = New CSocio
                 
                 If vsocio.LeerDatos(NuevoCliente) Then
                     sql2 = "insert into schfac (letraser, numfactu, fecfactu, codsocio, codcoope, codforpa, "
                     sql2 = sql2 & "baseimp1, baseimp2, baseimp3, impoiva1, impoiva2, impoiva3, tipoiva1,"
                     sql2 = sql2 & "tipoiva2, tipoiva3, porciva1, porciva2, porciva3, totalfac, impuesto,"
                     sql2 = sql2 & "intconta) values (" & DBSet(numser, "T") & "," & DBSet(vCont.Contador, "N") & ","
                     sql2 = sql2 & DBSet(NuevaFecFactu, "F") & "," & DBSet(NuevoCliente, "N") & "," & DBSet(vsocio.Colectivo, "N") & ","
                     sql2 = sql2 & DBSet(vsocio.ForPago, "N") & ","
                     sql2 = sql2 & DBSet(Rs!baseimp1, "N") & ","
                     sql2 = sql2 & DBSet(Rs!baseimp2, "N") & ","
                     sql2 = sql2 & DBSet(Rs!baseimp3, "N") & ","
                     sql2 = sql2 & DBSet(Rs!impoiva1, "N") & ","
                     sql2 = sql2 & DBSet(Rs!impoiva2, "N") & ","
                     sql2 = sql2 & DBSet(Rs!impoiva3, "N") & ","
                     sql2 = sql2 & DBSet(Rs!TipoIVA1, "N") & ","
                     sql2 = sql2 & DBSet(Rs!TipoIVA2, "N") & ","
                     sql2 = sql2 & DBSet(Rs!TipoIVA3, "N") & ","
                     sql2 = sql2 & DBSet(Rs!porciva1, "N") & ","
                     sql2 = sql2 & DBSet(Rs!porciva2, "N") & ","
                     sql2 = sql2 & DBSet(Rs!porciva3, "N") & ","
                     sql2 = sql2 & DBSet(Rs!TotalFac, "N") & ","
                     sql2 = sql2 & DBSet(Rs!Impuesto, "N") & ","
                     sql2 = sql2 & DBSet(0, "N") & ")"
        '             sql2 = sql2 & DBSet(RS!intconta, "N") & ")"
                 
                     caderr = "Insertando cabecera de factura de nuevo cliente:"
                     
                     NumError = db.ejecutar2(sql2, caderr)
                     
                     Rs.Close
                     
                     SQL = "select * from slhfac where letraser = " & DBSet(letraser, "T") & " and numfactu = " & DBSet(numfactu, "N")
                     SQL = SQL & " and fecfactu = " & DBSet(fecfactu, "F")
                     
            '         Set RS = New adodb.Recordset
            '         RS.Open sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                     Set Rs = db.cursor(SQL)
                     
                     caderr = "Insertando lineas de factura de nuevo cliente:"
                     
                     While Not Rs.EOF And (NumError = 0)
                        sql2 = "insert into slhfac (letraser,numfactu,fecfactu,numlinea,numalbar,fecalbar,horalbar,"
                        sql2 = sql2 & "codturno,numtarje,codartic,cantidad,preciove,implinea) values ("
                        sql2 = sql2 & DBSet(numser, "T") & "," & DBSet(vCont.Contador, "N") & "," & DBSet(NuevaFecFactu, "F") & ","
                        sql2 = sql2 & DBSet(Rs!NumLinea, "N") & "," & DBSet(Rs!numalbar, "T") & "," & DBSet(Rs!fecAlbar, "F") & ","
                        sql2 = sql2 & DBSet(Rs!horalbar, "FH") & "," & DBSet(Rs!codTurno, "N") & "," & DBSet(Rs!Numtarje, "N") & ","
                        sql2 = sql2 & DBSet(Rs!codartic, "N") & "," & DBSet(Rs!cantidad, "N") & ","
                        sql2 = sql2 & DBSet(Rs!preciove, "N") & "," & DBSet(Rs!ImpLinea, "N") & ")"
                        
                        NumError = db.ejecutar2(sql2, caderr)
                        
                        Rs.MoveNext
                     Wend
                End If
            End If
        End If
    End If
    
eCrearFacturaRectificativa:
    If Err.Number <> 0 Or NumError Then 'Or Not b Then
        If Err.Number <> 0 Then
            CrearFacturaRectificativa = Err.Number
        Else
            CrearFacturaRectificativa = NumError
        End If
        db.RollbackTrans
        ConnConta.RollbackTrans
    Else
        CrearFacturaRectificativa = 0
        db.CommitTrans
        ConnConta.CommitTrans
        
        
    End If
    Set db = Nothing
End Function

Public Function EsFacturaRectificable(LetraSerie As String) As Boolean
Dim SQL As String
    EsFacturaRectificable = False
    
    SQL = ""
    SQL = DevuelveDesdeBDNew(cPTours, "stipom", "letraser", "codtipom", "FAG", "T")
    
    EsFacturaRectificable = (Trim(SQL) = Trim(LetraSerie))
    
End Function


Public Function Prefacturacion(db As BaseDatos, DesFec As String, HasFec As String, dessoc As String, hassoc As String, descop As String, hascop As String, TipoClien As String) As Long
' funcion que cambia las formas de pago de los albaranes que sean distintos de contado y pone la forma de pago
' del cliente si ssocio.facturafp = 1
Dim SQL As String
Dim sql2 As String

Dim Rs As ADODB.Recordset

Dim Impuesto As Currency
Dim impbase As Currency
Dim ActSocio As Long
Dim ActForpa As Integer
Dim ActTarje As Long
Dim AntAlbaran As String
Dim AntTarje As Long
Dim AntSocio As Long
Dim AntForpa As Integer
Dim AntTurno As Integer
Dim HayReg As Boolean
Dim v_linea As Integer
Dim FamArtDto As String
Dim IvaArtDto As String
Dim ImporDto As Currency
Dim vCont As CContador
Dim DtoLitro As Currency
Dim CantCombustible As Currency
Dim Codigo As String
Dim baseimpo As Dictionary

Dim NumError As Long


    On Error GoTo ePrefacturacion

    
    SQL = "select scaalb.codclave, scaalb.codsocio, scaalb.codartic, scaalb.cantidad, scaalb.numlinea,"
    SQL = SQL & " scaalb.preciove, scaalb.importel, scaalb.numalbar, scaalb.fecalbar,"
    SQL = SQL & " scaalb.horalbar, scaalb.codturno, scaalb.codforpa, scaalb.numtarje "
    SQL = SQL & " from (((scaalb inner join ssocio on scaalb.codsocio = ssocio.codsocio) "
    SQL = SQL & " inner join scoope on ssocio.codcoope = scoope.codcoope and ssocio.facturafp = 1 "
    If descop <> "" Then SQL = SQL & " and ssocio.codcoope >= " & DBSet(descop, "N")
    If hascop <> "" Then SQL = SQL & " and ssocio.codcoope <= " & DBSet(hascop, "N")
    '[Monica]04/01/2013: efectivos
    SQL = SQL & ") inner join sforpa on scaalb.codforpa = sforpa.codforpa and sforpa.tipforpa <> 0 and sforpa.tipforpa <> 6) "
    SQL = SQL & " where scaalb.numfactu = 0 and scaalb.codforpa <> 98 "
    '[Monica]07/03/2018: solo los articulos q se facturen
    SQL = SQL & " and scaalb.codartic in (select codartic from sartic where facturar = 1) "
    If DesFec <> "" Then SQL = SQL & " and scaalb.fecalbar >= '" & Format(CDate(DesFec), FormatoFecha) & "' "
    If HasFec <> "" Then SQL = SQL & " and scaalb.fecalbar <= '" & Format(CDate(HasFec), FormatoFecha) & "' "
    If dessoc <> "" Then SQL = SQL & " and scaalb.codsocio >= " & DBSet(dessoc, "N")
    If hassoc <> "" Then SQL = SQL & " and scaalb.codsocio <= " & DBSet(hassoc, "N")
    
    Select Case TipoClien
        Case "0"
        
        Case "1"
            SQL = SQL & " and ssocio.bonifesp = 1"
        Case "2"
            SQL = SQL & " and ssocio.bonifesp = 0"
    End Select
    
    
    Set Rs = db.cursor(SQL)
    HayReg = False
    v_linea = 0
    NumError = 0
    While Not Rs.EOF And NumError = 0
        Forpa = DevuelveDesdeBDNew(cPTours, "ssocio", "codforpa", "codsocio", DBLet(Rs!codsocio, "N"), "N")
        sql2 = " update scaalb set codforpa = " & DBSet(Forpa, "N")
        sql2 = sql2 & " where codclave = " & DBSet(Rs!Codclave, "N")
        
        NumError = db.ejecutar(sql2)
        Rs.MoveNext
    Wend

ePrefacturacion:
    Prefacturacion = NumError
    Exit Function
End Function


Public Function ComprobarFechaVenci(FechaVenci As Date, Dia1 As Byte, Dia2 As Byte, Dia3 As Byte) As Date
Dim newFecha As Date
Dim b As Boolean

'=== Modificada Laura: 23/01/2007
    On Error GoTo ErrObtFec
    b = False
    
    '--- comprobar que tiene dias de pago para obtener nueva fecha
    If Not (Dia1 > 0 Or Dia2 > 0 Or Dia3 > 0) Then
        'si no tiene dias de pago la fecha es OK y fin
        ComprobarFechaVenci = FechaVenci
        Exit Function
    End If
        
    
    '--- Obtener nueva fecha del vencimiento
    newFecha = FechaVenci
    
    Do
        'si dia de la fecha vencimiento es uno de los 3 dias de pagos fecha es OK
        If Day(newFecha) = Dia1 Or Day(newFecha) = Dia2 Or Day(newFecha) = Dia3 Then
'            newFecha = CStr(newFecha)
            b = True
        Else
            'mientras esta en el mismo mes vamos aumentando dias hasta encontrar un dia de pago
            newFecha = DateAdd("d", 1, CDate(newFecha))
        End If
    Loop Until b = True Or Year(newFecha) = Year(FechaVenci) + 3
    
    ComprobarFechaVenci = newFecha
    Exit Function
    
ErrObtFec:
    MuestraError Err.Number, "Obtener Fecha vencimiento seg�n dias de pago.", Err.Description
End Function

Public Function ComprobarMesNoGira(FecVenci As Date, MesNG As Byte, DiaVtoAt As Byte, Dia1 As Byte, Dia2 As Byte, Dia3 As Byte) As Date
Dim F As String

    If Month(FecVenci) = MesNG Then
        If DiaVtoAt > 0 Then
            F = DiaVtoAt & "/"
        Else
            F = Day(FecVenci) & "/"
        End If
        
        If Month(FecVenci) + 1 < 13 Then
            F = F & Month(FecVenci) + 1 & "/" & Year(FecVenci)
        Else
            F = F & "01/" & Year(FecVenci) + 1
        End If
        FecVenci = Format(F, "dd/mm/yyyy")
    End If
    ComprobarMesNoGira = FecVenci
End Function

Public Function ModificacionAlbaranes(cadWhere As String, tabla As String, Pb1 As ProgressBar, Label4 As Label) As Boolean
Dim SQL As String
Dim sql2 As String
Dim Sql4 As String
Dim Rs As ADODB.Recordset
Dim rs4 As ADODB.Recordset
Dim margen As Currency
Dim EurosLitro As Double
Dim PrecioNue As Double
Dim PrecioNue2 As Double
Dim ImporteNue As Currency
Dim NRegs As Integer
Dim RsAlb As ADODB.Recordset
Dim CadenaAlb As String

    On Error GoTo eModificacionAlbaranes

    ModificacionAlbaranes = False

    Conn.BeginTrans

    CadenaAlb = ""
    Set RsAlb = New ADODB.Recordset
    RsAlb.Open Replace(Replace(cadWhere, "{", ""), "}", ""), Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not RsAlb.EOF
        CadenaAlb = CadenaAlb & DBLet(RsAlb!Codclave, "N") & ","
        RsAlb.MoveNext
    Wend
    Set RsAlb = Nothing
    If CadenaAlb <> "" Then CadenaAlb = Mid(CadenaAlb, 1, Len(CadenaAlb) - 1)
    
    '[Monica]07/03/2012: cambio del calculo para guardarnos el precio
    'Sql = "select scaalb.codclave, scaalb.codsocio, scaalb.codartic, scaalb.cantidad, tmpinformes.precio2 "
    SQL = "select distinct " & tabla & ".codsocio, " & tabla & ".codartic, tmpinformes.precio2 "
    SQL = SQL & " from " & tabla & " INNER JOIN tmpinformes ON " & tabla & ".codartic = tmpinformes.codigo1 and tmpinformes.codusu = " & vSesion.Codigo
    SQL = SQL & " where " & tabla & ".codclave in (" & CadenaAlb & ")" ' & Replace(Replace(cadWhere, "{", ""), "}", "") & ")"
    
    NRegs = TotalRegistrosConsulta(SQL)
    
    CargarProgres Pb1, NRegs
    Pb1.visible = True
    Label4.visible = True
    DoEvents
    
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    While Not Rs.EOF
        IncrementarProgres Pb1, 1
        DoEvents
        
        margen = DevuelveValor("select margen from smargen where codsocio = " & DBSet(Rs!codsocio, "N") & " and codartic = " & DBSet(Rs!codartic, "N"))
        '[Monica]15/12/2011: Euros/litro
        EurosLitro = DevuelveValor("select euroslitro from smargen where codsocio = " & DBSet(Rs!codsocio, "N") & " and codartic = " & DBSet(Rs!codartic, "N"))

        If margen <> 0 Then
            PrecioNue = CDbl(DBLet(Rs!precio2, "N")) * (1 + (margen / 100))
        Else
            PrecioNue = CDbl(DBLet(Rs!precio2, "N")) + EurosLitro
        End If
        
        PrecioNue2 = Round2(PrecioNue, 3)
        
        Sql4 = "select " & tabla & ".codclave, " & tabla & ".codsocio, " & tabla & ".codartic, " & tabla & ".cantidad, tmpinformes.precio2"
        Sql4 = Sql4 & " from " & tabla & " INNER JOIN tmpinformes ON " & tabla & ".codartic = tmpinformes.codigo1 and tmpinformes.codusu = " & vSesion.Codigo
        Sql4 = Sql4 & " where " & tabla & ".codclave in (" & CadenaAlb & ")" '& Replace(Replace(cadWhere, "{", ""), "}", "") & ")"
        Sql4 = Sql4 & " and " & tabla & ".codsocio = " & DBSet(Rs!codsocio, "N")
        Sql4 = Sql4 & " and " & tabla & ".codartic = " & DBSet(Rs!codartic, "N")
        
        Set rs4 = New ADODB.Recordset
        rs4.Open Sql4, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
        While Not rs4.EOF
            ImporteNue = Round2(PrecioNue * DBLet(rs4!cantidad, "N"), 2)
        
            '[Monica]15/12/2011: Precioinicio
            sql2 = "update " & tabla & " set precioinicial = preciove "
            sql2 = sql2 & " where codclave = " & DBSet(rs4!Codclave, "N")
            
            Conn.Execute sql2
    
            sql2 = "update " & tabla & " set preciove = " & DBSet(PrecioNue2, "N")
            sql2 = sql2 & " ,importel = " & DBSet(ImporteNue, "N")
            sql2 = sql2 & " where codclave = " & DBSet(rs4!Codclave, "N")
            
            Conn.Execute sql2
        
            rs4.MoveNext
        Wend
        Set rs4 = Nothing
        
        Rs.MoveNext
    Wend

    Set Rs = Nothing

    ModificacionAlbaranes = True
    Conn.CommitTrans
    Pb1.visible = False
    Label4.visible = False
    DoEvents
    Exit Function

eModificacionAlbaranes:
    Conn.RollbackTrans
    Pb1.visible = False
    Label4.visible = False
    DoEvents
    MuestraError Err.Number, "Modificacion Albaranes", Err.Description
End Function


'#####################################
Public Function SimulacionFacturacion(CliTar As Byte, Pb1 As ProgressBar, Label4 As Label, TipoGasoB As Byte) As Long
Dim SQL As String
Dim Rs As ADODB.Recordset

Dim Impuesto As Currency
Dim impbase As Currency
Dim ActSocio As Long
Dim ActForpa As Integer
Dim ActTarje As Long
Dim AntAlbaran As String
Dim AntTarje As Long
Dim AntSocio As Long
Dim AntForpa As Integer
Dim AntTurno As Integer
Dim HayReg As Boolean
Dim v_linea As Integer
Dim FamArtDto As String
Dim IvaArtDto As String
Dim ImporDto As Currency
Dim vCont As CContador
Dim DtoLitro As Currency
Dim CantCombustible As Currency
Dim Codigo As String

Dim NumError As Long
Dim numFac As Long

Dim tipoMov As String
Dim NRegs As Integer
Dim SqlAct As String
Dim TipForpa As String

    On Error GoTo eFacturacion

    SimulacionFacturacion = False


    Conn.Execute " DROP TABLE IF EXISTS tmpscaalb1;"
    SQL = "CREATE TEMPORARY TABLE tmpscaalb1 ( "
    SQL = SQL & "   `codsocio` mediumint(7) unsigned NOT NULL default '0',"
    SQL = SQL & "   `numfactu` int(7) NOT NULL default '0',"
    SQL = SQL & "   `numalbar` varchar(8) NOT NULL default '',"
    SQL = SQL & "   `fecalbar` date NOT NULL default '0000-00-00',"
    SQL = SQL & "   `horalbar` datetime NOT NULL default '0000-00-00 00:00:00',"
    SQL = SQL & "   `codturno` smallint(1) NOT NULL default '0',"
    SQL = SQL & "   `numtarje` int(8) NOT NULL default '0',"
    SQL = SQL & "   `codartic` int(6) NOT NULL default '0',"
    SQL = SQL & "   `cantidad` decimal(10,2) NOT NULL default '0.00',"
    SQL = SQL & "   `preciove` decimal(10,3) NOT NULL default '0.000',"
    SQL = SQL & "   `implinea` decimal(10,2) NOT NULL default '0.00', "
    SQL = SQL & "   KEY `codusu1` (`codsocio`)); "
    Conn.Execute SQL
    
    FamArtDto = "codfamia"
    IvaArtDto = DevuelveDesdeBD("codigiva", "sartic", "codartic", vParamAplic.ArticDto, "N", FamArtDto)
    
    SQL = "select tmpscaalb.codclave, tmpscaalb.codsocio, tmpscaalb.codartic, tmpscaalb.cantidad, tmpscaalb.numlinea,"
    SQL = SQL & " tmpscaalb.preciove, tmpscaalb.importel, tmpscaalb.numalbar, tmpscaalb.fecalbar,"
    SQL = SQL & " tmpscaalb.horalbar, tmpscaalb.codturno, tmpscaalb.codforpa, tmpscaalb.numtarje, tmpscaalb.matricul, tmpscaalb.precioinicial "
    SQL = SQL & " from ((tmpscaalb inner join ssocio on tmpscaalb.codsocio = ssocio.codsocio) "
    SQL = SQL & " inner join scoope on ssocio.codcoope = scoope.codcoope) "

    
    If vParamAplic.Cooperativa = 4 Then
        SQL = SQL & " inner join sartic on tmpscaalb.codartic = sartic.codartic and sartic.tipogaso = 0 "
    End If
    
    SQL = SQL & " where codusu = " & vSesion.Codigo
    
    '[Monica]19/06/2013: A�adimos el if de cooperativa y tipogasob
    If (vParamAplic.Cooperativa = 1 Or vParamAplic.Cooperativa = 2) And TipoGasoB > 0 Then
        ' no miramos si es por cliente o por tarjeta
        
        '[Monica]15/07/2013: a�adido el caso de que sea interna
        If CliTar = 3 Then
            SQL = SQL & " and scoope.tipfactu = " & DBLet(CliTar, "N")
        Else
            ' no miramos si es por cliente o por tarjeta
'            Sql = Sql & " and scoope.tipfactu <= " & DBLet(CliTar, "N")
            SQL = SQL & " and scoope.tipfactu in (0,1) "
        End If
        
    Else
        SQL = SQL & " and scoope.tipfactu = " & DBLet(CliTar, "N")
    End If
    
    
    '[Monica]28/07/2011: en el caso de las facturas internas quieren que sea por tarjeta antes era por cliente
    If CliTar = 1 Then
        SQL = SQL & " order by tmpscaalb.codsocio, tmpscaalb.codforpa, tmpscaalb.fecalbar, tmpscaalb.numalbar, tmpscaalb.numlinea, tmpscaalb.codclave "
    Else
        SQL = SQL & " order by tmpscaalb.codsocio, tmpscaalb.numtarje, tmpscaalb.codforpa, tmpscaalb.fecalbar, tmpscaalb.numalbar, tmpscaalb.numlinea, tmpscaalb.codclave "
    End If
    
    
    NRegs = TotalRegistrosConsulta(SQL)
    CargarProgres Pb1, NRegs
    Pb1.visible = True
    Label4.visible = True
    Label4.Caption = "Simulando Facturacion:"
    DoEvents
    
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    HayReg = False
    v_linea = 0
    NumError = 0
    If Not Rs.EOF Then
        Rs.MoveFirst
        AntSocio = Rs!codsocio
        AntAlbaran = Rs!numalbar
        AntForpa = Rs!Codforpa
        AntTurno = Rs!codTurno
        AntTarje = Rs!Numtarje
        
        numFac = 1
        
        TotalImp = 0
        TotalImpSigaus = 0
        
        b = True
        
        While Not Rs.EOF And b
            
            IncrementarProgres Pb1, 1
            Label4.Caption = "Simulando Facturacion: " & Format(Rs!codsocio, "000000") & " - " & Rs!numalbar
            DoEvents
            
            
            HayReg = True
            ActForpa = Rs!Codforpa
            ActSocio = Rs!codsocio
            ActTarje = Rs!Numtarje
            If ((ActForpa <> AntForpa Or ActSocio <> AntSocio) And (CliTar = 1 Or CliTar = 3)) Or _
            ((ActForpa <> AntForpa Or ActSocio <> AntSocio Or ActTarje <> AntTarje) And (CliTar = 0 Or (CliTar = 3 And TipoGasoB <> 0))) Then ' after group of codforpa
            
               '  ### [Monica] 05/12/2006
               ' modificacion: si la forma de pago no admite bonificacion no hacemos
               If AdmiteBonificacion(AntForpa) Then
 
                   ' miramos el descuento/litro de socio sobre la cantidad
                   SQL = ""
                   SQL = DevuelveDesdeBD("dtolitro", "ssocio", "codsocio", CStr(AntSocio), "N")
                   DtoLitro = 0
                   If SQL <> "" Then DtoLitro = CCur(SQL)
    
                   If DtoLitro <> 0 And b Then
                        DtoLitro = DtoLitro * (-1)
                        ImporDto = Round2(CantCombustible * DtoLitro, 2)
                        b = InsertaLineaDescuentoSimula(numFac, AntSocio, CantCombustible, ImporDto, DtoLitro, AntTarje)
                   End If
               End If
               
               If b Then
                    numFac = numFac + 1
               End If

               TotalImp = 0
               TotalImpSigaus = 0
               AntForpa = ActForpa
               AntSocio = ActSocio
               AntTurno = Rs!codTurno
               AntTarje = ActTarje
               
               CantCombustible = 0
               
               ImpFactu = 0
            End If
            
            '[Monica]24/01/2013: si el socio es un cliente no de varios vemos si hay q partirle la factura
            TipForpa = DevuelveDesdeBDNew(cPTours, "sforpa", "tipforpa", "codforpa", CStr(AntForpa), "N")
            If vParamAplic.Cooperativa = 1 And Not EsDeVarios(CStr(AntSocio)) And vParamAplic.LimiteFra <> 0 And (ImpFactu + DBLet(Rs!importel, "N") > vParamAplic.LimiteFra) And TipForpa = "0" Then
               
                If b Then
                    numFac = numFac + 1
                End If

                TotalImp = 0
                TotalImpSigaus = 0
               
                CantCombustible = 0
               
                ImpFactu = 0
            
            Else
                '[Monica]24/01/2013: sumamos el total factura
                ImpFactu = ImpFactu + DBLet(Rs!importel, "N")
                
                '-------
                ' tenemos que calcular el impuesto multiplicando cantidad de linea por impuesto por articulo
                Codigo = "codigiva"
                Sql1 = ""
                Sql1 = DevuelveDesdeBD("impuesto", "sartic", "codartic", DBLet(Rs!codartic), "N", Codigo)
                If Sql1 = "" Then
                    Impuesto = 0
                Else
                    Impuesto = CCur(Sql1) ' Comprueba si es nulo y lo pone a 0 o ""
                End If
                
                If EsArticuloCombustible(Rs!codartic) Then
                    TotalImp = TotalImp + Round2((Rs!cantidad * Impuesto), 2)
                    CantCombustible = CantCombustible + DBLet(Rs!cantidad, "N")
                End If
                
                '[Monica]15/02/2011: cuando el articulo es lubricante, tiene un impuesto, hemos de calcularlo
                ' Sabemos que es lubricante pq tiene un peso por unidad.
                ' El Impuesto se calcula multiplicandolo por el preciosigaus
                TotalImpSigaus = TotalImpSigaus + ImpuestoSigausArticulo(Rs!codartic, Rs!cantidad)
                
                SqlAct = "update tmpscaalb set numfactu = " & DBSet(numFac, "N") & " where codusu = " & vSesion.Codigo & " and codclave = " & DBSet(Rs!Codclave, "N")
                Conn.Execute SqlAct
                
                Rs.MoveNext
            
            End If
        Wend
        If HayReg And b Then
            ' miramos el descuento/litro de socio sobre la cantidad
            If AdmiteBonificacion(AntForpa) Then
                 SQL = ""
                 SQL = DevuelveDesdeBD("dtolitro", "ssocio", "codsocio", CStr(AntSocio), "N")
                 DtoLitro = 0
                 If SQL <> "" Then DtoLitro = CCur(SQL)
                 If DtoLitro <> 0 And b Then
                      DtoLitro = DtoLitro * (-1)
                      ImporDto = Round2(CantCombustible * DtoLitro, 2)
                      b = InsertaLineaDescuentoSimula(numFac, AntSocio, CantCombustible, ImporDto, DtoLitro, AntTarje)
                 End If
            End If
            
            ' por ultimo insertamos las bonificaciones de la tabla temporal en tmpscaalb1
            SqlAct = "insert into tmpscaalb (codusu, codsocio, numfactu, numalbar, fecalbar, horalbar, codturno, numtarje, codartic, cantidad, preciove, importel)"
            SqlAct = SqlAct & " select " & vSesion.Codigo & ", codsocio, numfactu, numalbar, fecalbar, horalbar, codturno, numtarje, codartic, cantidad, preciove, implinea "
            SqlAct = SqlAct & " from tmpscaalb1 "
            Conn.Execute SqlAct
            
            
        End If
    End If
    
    Conn.Execute " DROP TABLE IF EXISTS tmpscaalb1;"
    
    SimulacionFacturacion = True
    Pb1.visible = False
    Label4.visible = False
    Exit Function
    
eFacturacion:
    MensError = Err.Description
End Function

Public Function InsertaLineaDescuentoSimula(numFac As Long, socio As Long, cantidad As Currency, Importe As Currency, Precio As Currency, Tarjeta As Long) As Boolean
    Dim SQL As String
    Dim ImpLinea As Currency
    Dim Texto As String
    
    '05012007
    On Error GoTo eInsertaLineaDescuentoSimula
    
    InsertaLineaDescuentoSimula = False
    
    MensError = ""
    Texto = "BONIFICA"
    
    SQL = "INSERT into tmpscaalb1 "
    
    SQL = SQL & "(codsocio, numfactu, numalbar, " & _
           "fecalbar, horalbar, codturno, numtarje, codartic, " & _
           "cantidad, preciove, implinea) " & _
           "values " & _
           "(" & DBSet(socio, "N") & "," & DBSet(numFac, "N") & "," & DBSet(Texto, "T") & "," & _
           DBSet(Now, "F") & "," & DBSet(Now, "FH") & ",1," & DBSet(Tarjeta, "N") & "," & DBSet(vParamAplic.ArticDto, "N") & "," & _
           DBSet(cantidad, "N") & "," & DBSet(Precio, "N") & "," & DBSet(Importe, "N") & ")"
    
    Conn.Execute SQL
           
    InsertaLineaDescuentoSimula = True
    Exit Function
    
eInsertaLineaDescuentoSimula:
    MensError = "Error insertando linea descuento bonificacion"
End Function


Public Function EsInterna(Letra As String) As Boolean
Dim SQL As String

    SQL = "select esinterna from stipom where letraser = '" & Trim(Letra) & "'"
    
    EsInterna = (DevuelveValor(SQL) = 1)

End Function
