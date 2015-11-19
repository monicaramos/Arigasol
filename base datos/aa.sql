update ssocio set grupoestartic = 4 where codsocio >= 10000 and codsocio <= 10999



SELECT sum(cantidad), sum(implinea), slhfac.codartic, ssocio.grupoestartic , slhfac.fecalbar
 FROM   slhfac, schfac, ssocio, sfamia, sartic
 where slhfac.letraser = schfac.letraser and slhfac.numfactu = schfac.numfactu and 
       slhfac.fecfactu = schfac.fecfactu and slhfac.codartic = sartic.codartic and
       sartic.codfamia = sfamia.codfamia and schfac.codsocio = ssocio.codsocio
 group by  slhfac.codartic, ssocio.grupoestartic, slhfac.fecalbar
 ORDER BY `slhfac`.`codartic`, ssocio.grupoestartic, slhfac.fecalbar desc