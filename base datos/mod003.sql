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

/*Table structure for table `cambios` */

CREATE TABLE `cambios` (
  `codusu` smallint(3) unsigned NOT NULL default '0' COMMENT 'Codigo de usuario',
  `fechacambio` datetime NOT NULL COMMENT 'Fecha de Cambio',
  `tabla` varchar(30) NOT NULL COMMENT 'Tabla sobre la que se hace cambio',
  `cadena` varchar(500) NOT NULL COMMENT 'Sql que ejecuta',
  `valoranterior` varchar(500) default NULL COMMENT 'Valores anteriores'
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

SET SQL_MODE=@OLD_SQL_MODE;
SET FOREIGN_KEY_CHECKS=@OLD_FOREIGN_KEY_CHECKS;
