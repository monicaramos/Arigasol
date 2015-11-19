/*
SQLyog - Free MySQL GUI v5.18
Host - 5.0.22 : Database - arigasol
*********************************************************************
Server version : 5.0.22
*/ 
SET NAMES utf8;

SET SQL_MODE='';

USE `arigasol`;

SET @OLD_FOREIGN_KEY_CHECKS=@@FOREIGN_KEY_CHECKS, FOREIGN_KEY_CHECKS=0;
SET @OLD_SQL_MODE=@@SQL_MODE, SQL_MODE='NO_AUTO_VALUE_ON_ZERO';

/*Table structure for table `scagru` */

CREATE TABLE `scagru` (
  `codempre` int(4) unsigned NOT NULL default '0' COMMENT 'Codigo',
  `nomempre` varchar(30) NOT NULL COMMENT 'Denominacion',
  PRIMARY KEY  (`codempre`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

/*Table structure for table `sligru` */

CREATE TABLE `sligru` (
  `codempre` int(4) NOT NULL default '0' COMMENT 'Codigo',
  `codsocio` int(6) NOT NULL default '0' COMMENT 'Socio',
  PRIMARY KEY  (`codempre`,`codsocio`),
  KEY `FK_sligru` (`codsocio`),
  CONSTRAINT `sligru_ibfk_1` FOREIGN KEY (`codsocio`) REFERENCES `ssocio` (`codsocio`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

SET SQL_MODE=@OLD_SQL_MODE;
SET FOREIGN_KEY_CHECKS=@OLD_FOREIGN_KEY_CHECKS;


/* Nuevas tablas de historico de facturas para el Regaixo */
/*
SQLyog - Free MySQL GUI v5.18
Host - 5.0.22 : Database - arigasol
*********************************************************************
Server version : 5.0.22
*/ 
SET NAMES utf8;

SET SQL_MODE='';

USE `arigasol`;

SET @OLD_FOREIGN_KEY_CHECKS=@@FOREIGN_KEY_CHECKS, FOREIGN_KEY_CHECKS=0;
SET @OLD_SQL_MODE=@@SQL_MODE, SQL_MODE='NO_AUTO_VALUE_ON_ZERO';

/*Table structure for table `schfacr` */

CREATE TABLE `schfacr` (
  `letraser` char(1) NOT NULL default '',
  `numfactu` int(7) NOT NULL default '0',
  `fecfactu` date NOT NULL default '0000-00-00',
  `codsocio` int(6) unsigned NOT NULL default '0',
  `codcoope` smallint(2) unsigned NOT NULL default '0',
  `codforpa` smallint(2) unsigned NOT NULL default '0',
  `baseimp1` decimal(10,2) NOT NULL default '0.00',
  `baseimp2` decimal(10,2) default '0.00',
  `baseimp3` decimal(10,2) default '0.00',
  `impoiva1` decimal(10,2) NOT NULL default '0.00',
  `impoiva2` decimal(10,2) default '0.00',
  `impoiva3` decimal(10,2) default '0.00',
  `tipoiva1` smallint(2) NOT NULL default '0',
  `tipoiva2` smallint(2) default '0',
  `tipoiva3` smallint(2) default '0',
  `porciva1` decimal(4,2) NOT NULL default '0.00',
  `porciva2` decimal(4,2) default '0.00',
  `porciva3` decimal(4,2) default '0.00',
  `totalfac` decimal(10,2) NOT NULL default '0.00',
  `impuesto` decimal(10,2) NOT NULL default '0.00',
  `intconta` tinyint(1) NOT NULL default '0',
  PRIMARY KEY  (`letraser`,`numfactu`,`fecfactu`),
  KEY `codcoope` (`codcoope`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

/*Table structure for table `slhfacr` */

CREATE TABLE `slhfacr` (
  `letraser` char(1) NOT NULL default '',
  `numfactu` int(7) unsigned NOT NULL default '0',
  `fecfactu` date NOT NULL default '0000-00-00',
  `numlinea` smallint(4) NOT NULL default '0',
  `numalbar` varchar(8) NOT NULL default '',
  `fecalbar` date NOT NULL default '0000-00-00',
  `horalbar` datetime NOT NULL default '0000-00-00 00:00:00',
  `codturno` smallint(1) NOT NULL default '0',
  `numtarje` int(8) NOT NULL default '0',
  `codartic` int(6) NOT NULL default '0',
  `cantidad` decimal(10,2) NOT NULL default '0.00',
  `preciove` decimal(10,3) NOT NULL default '0.000',
  `implinea` decimal(10,2) NOT NULL default '0.00',
  PRIMARY KEY  (`letraser`,`numfactu`,`fecfactu`,`numlinea`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

SET SQL_MODE=@OLD_SQL_MODE;
SET FOREIGN_KEY_CHECKS=@OLD_FOREIGN_KEY_CHECKS;

