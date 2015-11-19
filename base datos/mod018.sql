/* Tabla temporal para la reimpresion de facturas, por caso de querer reimprimir facturas que tengan 
algún artículo de gasoleo B */

USE `arigasol`;

SET @OLD_FOREIGN_KEY_CHECKS=@@FOREIGN_KEY_CHECKS, FOREIGN_KEY_CHECKS=0;
SET @OLD_SQL_MODE=@@SQL_MODE, SQL_MODE='NO_AUTO_VALUE_ON_ZERO';

/*Table structure for table `tmpfacturas` */

CREATE TABLE `tmpfacturas` (
  `codusu` smallint(3),
  `letraser` char(1) default NULL,
  `numfactu` int(7) default NULL,
  `fecfactu` date default NULL
) ENGINE=MyISAM DEFAULT CHARSET=latin1 COMMENT='Temporal para reimpresion de facturas';

alter table `tmpfacturas` add unique `letraser` ( `letraser`, `numfactu`, `fecfactu` )

SET SQL_MODE=@OLD_SQL_MODE;
SET FOREIGN_KEY_CHECKS=@OLD_FOREIGN_KEY_CHECKS;
