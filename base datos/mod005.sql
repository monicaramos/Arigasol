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

/*Table structure for table `soltarje` */

CREATE TABLE `soltarje` (
  `codsocio` int(6) NOT NULL default '0' COMMENT 'Socio',
  `tipotarje` smallint(1) NOT NULL default '0' COMMENT 'tipo:0=normal 1=bonificada',
  `numtarje` int(3) NOT NULL default '0' COMMENT 'Num.Tarjetas',
  PRIMARY KEY  (`codsocio`,`tipotarje`),
  CONSTRAINT `soltarje_ibfk_1` FOREIGN KEY (`codsocio`) REFERENCES `ssocio` (`codsocio`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

SET SQL_MODE=@OLD_SQL_MODE;
SET FOREIGN_KEY_CHECKS=@OLD_FOREIGN_KEY_CHECKS;

