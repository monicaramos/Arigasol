USE `arigasol`;

SET @OLD_FOREIGN_KEY_CHECKS=@@FOREIGN_KEY_CHECKS, FOREIGN_KEY_CHECKS=0;
SET @OLD_SQL_MODE=@@SQL_MODE, SQL_MODE='NO_AUTO_VALUE_ON_ZERO';

/*Table structure for table `gp_suministro` */

CREATE TABLE `gp_suministro` (
  `id` varchar(20) NOT NULL,
  `idmovcont` int(11) NOT NULL default '0',
  `cim` varchar(13) NOT NULL,
  `fechahora` datetime NOT NULL,
  `codprod` smallint(6) NOT NULL default '0' COMMENT 'Porcentaje biodiesel 0-100',
  `lit` decimal(7,2) NOT NULL,
  `nif` varchar(9) NOT NULL,
  `matricula` varchar(12) NOT NULL,
  PRIMARY KEY  (`id`,`idmovcont`),
  CONSTRAINT `gp_suministro_ibfk_1` FOREIGN KEY (`id`) REFERENCES `gp_suministrv2ent` (`id`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1 ROW_FORMAT=DYNAMIC;

/*Table structure for table `gp_suministrv2ent` */

CREATE TABLE `gp_suministrv2ent` (
  `id` varchar(20) NOT NULL,
  `codee` varchar(4) default NULL,
  `test` varchar(1) NOT NULL,
  `situacion` smallint(6) default '0' COMMENT '0 = pendiente / 1 = enviado',
  PRIMARY KEY  (`id`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1 ROW_FORMAT=DYNAMIC;

SET SQL_MODE=@OLD_SQL_MODE;
SET FOREIGN_KEY_CHECKS=@OLD_FOREIGN_KEY_CHECKS;





