/*
SQLyog - Free MySQL GUI v5.18
Host - 5.0.22 : Database - arigasol
*********************************************************************
Server version : 5.0.22
*/ 
SET NAMES utf8;

SET SQL_MODE='';

create database if not exists `arigasol`;

USE `arigasol`;

SET @OLD_FOREIGN_KEY_CHECKS=@@FOREIGN_KEY_CHECKS, FOREIGN_KEY_CHECKS=0;
SET @OLD_SQL_MODE=@@SQL_MODE, SQL_MODE='NO_AUTO_VALUE_ON_ZERO';

/*Table structure for table `appmenus` */

CREATE TABLE `appmenus` (
  `aplicacion` varchar(15) default '0',
  `Name` varchar(100) default '0',
  `caption` varchar(100) default '0',
  `indice` tinyint(3) default '0',
  `padre` smallint(3) unsigned default '0',
  `orden` tinyint(3) unsigned default NULL,
  `Contador` smallint(5) unsigned default NULL
) ENGINE=MyISAM DEFAULT CHARSET=latin1;

/*Data for the table `appmenus` */

insert into `appmenus` (`aplicacion`,`Name`,`caption`,`indice`,`padre`,`orden`,`Contador`) values ('Arigasol','mnE_Tanques','Datos Tanques/Mangueras',1,45,1,46);
insert into `appmenus` (`aplicacion`,`Name`,`caption`,`indice`,`padre`,`orden`,`Contador`) values ('Arigasol','mnTanques','Tanques-Mangueras',NULL,0,5,45);
insert into `appmenus` (`aplicacion`,`Name`,`caption`,`indice`,`padre`,`orden`,`Contador`) values ('Arigasol','mnE_Estadist','Ventas Artículos por Tarjeta',7,37,7,44);
insert into `appmenus` (`aplicacion`,`Name`,`caption`,`indice`,`padre`,`orden`,`Contador`) values ('Arigasol','mnE_Estadist','Resumen Ventas Diarias',6,37,6,43);
insert into `appmenus` (`aplicacion`,`Name`,`caption`,`indice`,`padre`,`orden`,`Contador`) values ('Arigasol','mnE_Estadist','Resumen Ventas Artículos',5,37,5,42);
insert into `appmenus` (`aplicacion`,`Name`,`caption`,`indice`,`padre`,`orden`,`Contador`) values ('Arigasol','mnE_Estadist','Ventas Artículos por Cliente',4,37,4,41);
insert into `appmenus` (`aplicacion`,`Name`,`caption`,`indice`,`padre`,`orden`,`Contador`) values ('Arigasol','mnE_Estadist','Ventas por Cliente',3,37,3,40);
insert into `appmenus` (`aplicacion`,`Name`,`caption`,`indice`,`padre`,`orden`,`Contador`) values ('Arigasol','mnE_Estadist','Diario de Facturación',2,37,2,39);
insert into `appmenus` (`aplicacion`,`Name`,`caption`,`indice`,`padre`,`orden`,`Contador`) values ('Arigasol','mnE_Estadist','Historico Facturas',1,37,1,38);
insert into `appmenus` (`aplicacion`,`Name`,`caption`,`indice`,`padre`,`orden`,`Contador`) values ('Arigasol','mnEstadisticas','Estadísticas',NULL,0,4,37);
insert into `appmenus` (`aplicacion`,`Name`,`caption`,`indice`,`padre`,`orden`,`Contador`) values ('Arigasol','mnF_Facturacion','Contabilizar Facturación',5,31,5,36);
insert into `appmenus` (`aplicacion`,`Name`,`caption`,`indice`,`padre`,`orden`,`Contador`) values ('Arigasol','mnF_Facturacion','-',4,31,4,35);
insert into `appmenus` (`aplicacion`,`Name`,`caption`,`indice`,`padre`,`orden`,`Contador`) values ('Arigasol','mnF_Facturacion','Reimpresión de Facturas',3,31,3,34);
insert into `appmenus` (`aplicacion`,`Name`,`caption`,`indice`,`padre`,`orden`,`Contador`) values ('Arigasol','mnF_Facturacion','Facturación',2,31,2,33);
insert into `appmenus` (`aplicacion`,`Name`,`caption`,`indice`,`padre`,`orden`,`Contador`) values ('Arigasol','mnF_Facturacion','Traspaso Facturas Tpv',1,31,1,32);
insert into `appmenus` (`aplicacion`,`Name`,`caption`,`indice`,`padre`,`orden`,`Contador`) values ('Arigasol','mnFacturacion','Facturación',NULL,0,3,31);
insert into `appmenus` (`aplicacion`,`Name`,`caption`,`indice`,`padre`,`orden`,`Contador`) values ('Arigasol','mnG_Ventas','-',10,19,10,29);
insert into `appmenus` (`aplicacion`,`Name`,`caption`,`indice`,`padre`,`orden`,`Contador`) values ('Arigasol','mnG_Ventas','Contabilizar Cierre Turno',11,19,11,30);
insert into `appmenus` (`aplicacion`,`Name`,`caption`,`indice`,`padre`,`orden`,`Contador`) values ('Arigasol','mnG_Ventas','Cuadre diario',9,19,9,28);
insert into `appmenus` (`aplicacion`,`Name`,`caption`,`indice`,`padre`,`orden`,`Contador`) values ('Arigasol','mnG_Ventas','-',8,19,8,27);
insert into `appmenus` (`aplicacion`,`Name`,`caption`,`indice`,`padre`,`orden`,`Contador`) values ('Arigasol','mnG_Ventas','Comprobación descuadres',7,19,7,26);
insert into `appmenus` (`aplicacion`,`Name`,`caption`,`indice`,`padre`,`orden`,`Contador`) values ('Arigasol','mnG_Ventas','Buscar errores Albaranes',6,19,6,25);
insert into `appmenus` (`aplicacion`,`Name`,`caption`,`indice`,`padre`,`orden`,`Contador`) values ('Arigasol','mnG_Ventas','Informe Prefacturación',5,19,5,24);
insert into `appmenus` (`aplicacion`,`Name`,`caption`,`indice`,`padre`,`orden`,`Contador`) values ('Arigasol','mnG_Ventas','Resumen Ventas Articulos',4,19,4,23);
insert into `appmenus` (`aplicacion`,`Name`,`caption`,`indice`,`padre`,`orden`,`Contador`) values ('Arigasol','mnG_Ventas','Albaranes',3,19,3,22);
insert into `appmenus` (`aplicacion`,`Name`,`caption`,`indice`,`padre`,`orden`,`Contador`) values ('Arigasol','mnG_Ventas','-',2,19,2,21);
insert into `appmenus` (`aplicacion`,`Name`,`caption`,`indice`,`padre`,`orden`,`Contador`) values ('Arigasol','mnG_Ventas','Traspaso datos Postes',1,19,1,20);
insert into `appmenus` (`aplicacion`,`Name`,`caption`,`indice`,`padre`,`orden`,`Contador`) values ('Arigasol','mnGeneral','Ventas Diarias',NULL,0,2,19);
insert into `appmenus` (`aplicacion`,`Name`,`caption`,`indice`,`padre`,`orden`,`Contador`) values ('Arigasol','mnP_Generales','-',16,1,16,17);
insert into `appmenus` (`aplicacion`,`Name`,`caption`,`indice`,`padre`,`orden`,`Contador`) values ('Arigasol','mnP_Generales','Salir',17,1,17,18);
insert into `appmenus` (`aplicacion`,`Name`,`caption`,`indice`,`padre`,`orden`,`Contador`) values ('Arigasol','mnP_Generales','Bancos propios',15,1,15,16);
insert into `appmenus` (`aplicacion`,`Name`,`caption`,`indice`,`padre`,`orden`,`Contador`) values ('Arigasol','mnP_Generales','Formas de pago',14,1,14,15);
insert into `appmenus` (`aplicacion`,`Name`,`caption`,`indice`,`padre`,`orden`,`Contador`) values ('Arigasol','mnP_Generales','Situaciones',13,1,13,14);
insert into `appmenus` (`aplicacion`,`Name`,`caption`,`indice`,`padre`,`orden`,`Contador`) values ('Arigasol','mnP_Generales','Artículos',12,1,12,13);
insert into `appmenus` (`aplicacion`,`Name`,`caption`,`indice`,`padre`,`orden`,`Contador`) values ('Arigasol','mnP_Generales','Empleados',10,1,10,11);
insert into `appmenus` (`aplicacion`,`Name`,`caption`,`indice`,`padre`,`orden`,`Contador`) values ('Arigasol','mnP_Generales','Familias',11,1,11,12);
insert into `appmenus` (`aplicacion`,`Name`,`caption`,`indice`,`padre`,`orden`,`Contador`) values ('Arigasol','mnP_Generales','Clientes',9,1,9,10);
insert into `appmenus` (`aplicacion`,`Name`,`caption`,`indice`,`padre`,`orden`,`Contador`) values ('Arigasol','mnP_Generales','Colectivos',8,1,8,9);
insert into `appmenus` (`aplicacion`,`Name`,`caption`,`indice`,`padre`,`orden`,`Contador`) values ('Arigasol','mnP_Generales','-',7,1,7,8);
insert into `appmenus` (`aplicacion`,`Name`,`caption`,`indice`,`padre`,`orden`,`Contador`) values ('Arigasol','mnP_Generales','Usuarios',6,1,6,7);
insert into `appmenus` (`aplicacion`,`Name`,`caption`,`indice`,`padre`,`orden`,`Contador`) values ('Arigasol','mnP_Generales','Tipo de Crédito ',5,1,5,6);
insert into `appmenus` (`aplicacion`,`Name`,`caption`,`indice`,`padre`,`orden`,`Contador`) values ('Arigasol','mnP_Generales','Tipos de Documentos',4,1,4,5);
insert into `appmenus` (`aplicacion`,`Name`,`caption`,`indice`,`padre`,`orden`,`Contador`) values ('Arigasol','mnP_Generales','Tipos de Movimiento',3,1,3,4);
insert into `appmenus` (`aplicacion`,`Name`,`caption`,`indice`,`padre`,`orden`,`Contador`) values ('Arigasol','mnP_Generales','Parámetros',2,1,2,3);
insert into `appmenus` (`aplicacion`,`Name`,`caption`,`indice`,`padre`,`orden`,`Contador`) values ('Arigasol','mnP_Generales','Datos de Empresa',1,1,1,2);
insert into `appmenus` (`aplicacion`,`Name`,`caption`,`indice`,`padre`,`orden`,`Contador`) values ('Arigasol','mnParametros','Datos Básicos',1,0,1,1);
insert into `appmenus` (`aplicacion`,`Name`,`caption`,`indice`,`padre`,`orden`,`Contador`) values ('Arigasol','mnE_Tanques','Datos Recaudación',2,45,2,47);
insert into `appmenus` (`aplicacion`,`Name`,`caption`,`indice`,`padre`,`orden`,`Contador`) values ('Arigasol','mnUtil','Utilidades',NULL,0,6,48);
insert into `appmenus` (`aplicacion`,`Name`,`caption`,`indice`,`padre`,`orden`,`Contador`) values ('Arigasol','mnE_Util','Copia de Seguridad local',1,48,1,49);
insert into `appmenus` (`aplicacion`,`Name`,`caption`,`indice`,`padre`,`orden`,`Contador`) values ('Arigasol','mnSoporte','Soporte',NULL,0,7,50);
insert into `appmenus` (`aplicacion`,`Name`,`caption`,`indice`,`padre`,`orden`,`Contador`) values ('Arigasol','mnE_Soporte1','Web Soporte',NULL,50,1,51);
insert into `appmenus` (`aplicacion`,`Name`,`caption`,`indice`,`padre`,`orden`,`Contador`) values ('Arigasol','mnp_Barra2','-',NULL,50,2,52);
insert into `appmenus` (`aplicacion`,`Name`,`caption`,`indice`,`padre`,`orden`,`Contador`) values ('Arigasol','mnE_Soporte2','Acerca de',NULL,50,3,53);

/*Table structure for table `appmenususuario` */

CREATE TABLE `appmenususuario` (
  `aplicacion` varchar(15) NOT NULL default '0',
  `codusu` smallint(1) unsigned NOT NULL default '0',
  `codigo` smallint(3) unsigned NOT NULL default '0',
  `tag` varchar(100) default '0',
  PRIMARY KEY  (`aplicacion`,`codusu`,`codigo`)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;

/*Data for the table `appmenususuario` */

insert into `appmenususuario` (`aplicacion`,`codusu`,`codigo`,`tag`) values ('Arigasol',0,1,'mnAlmacen|');
insert into `appmenususuario` (`aplicacion`,`codusu`,`codigo`,`tag`) values ('Arigasol',1,1,'mnP_Generales|5');
insert into `appmenususuario` (`aplicacion`,`codusu`,`codigo`,`tag`) values ('Arigasol',1,2,'mnP_Generales|6');
insert into `appmenususuario` (`aplicacion`,`codusu`,`codigo`,`tag`) values ('Arigasol',2,1,'mnP_Generales|5');
insert into `appmenususuario` (`aplicacion`,`codusu`,`codigo`,`tag`) values ('Arigasol',2,2,'mnP_Generales|6');

/*Table structure for table `pcs` */

CREATE TABLE `pcs` (
  `codpc` smallint(5) unsigned NOT NULL default '0',
  `nompc` char(30) default NULL,
  `Conta` smallint(5) unsigned default NULL,
  PRIMARY KEY  (`codpc`)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;

/*Data for the table `pcs` */

insert into `pcs` (`codpc`,`nompc`,`Conta`) values (1,'PCDAVIDG',NULL);
insert into `pcs` (`codpc`,`nompc`,`Conta`) values (2,'PCLAURA',NULL);
insert into `pcs` (`codpc`,`nompc`,`Conta`) values (3,'ARIADNA',NULL);
insert into `pcs` (`codpc`,`nompc`,`Conta`) values (4,'RAFA',NULL);
insert into `pcs` (`codpc`,`nompc`,`Conta`) values (5,'MANOLO',NULL);
insert into `pcs` (`codpc`,`nompc`,`Conta`) values (6,'dhcppc2',NULL);
insert into `pcs` (`codpc`,`nompc`,`Conta`) values (7,'trinux',NULL);
insert into `pcs` (`codpc`,`nompc`,`Conta`) values (8,'ARIADNA2',NULL);
insert into `pcs` (`codpc`,`nompc`,`Conta`) values (9,'pcMonica',NULL);

/*Table structure for table `sartic` */

CREATE TABLE `sartic` (
  `codartic` int(6) NOT NULL default '0',
  `nomartic` varchar(40) NOT NULL default '',
  `codfamia` smallint(3) NOT NULL default '0',
  `codigean` varchar(13) default NULL COMMENT 'Null o unico.',
  `numtanqu` smallint(3) default NULL COMMENT 'Solo Familia Carburantes.',
  `nummangu` smallint(3) default NULL COMMENT 'Solo Familia Carburantes.',
  `codmacta` varchar(10) default NULL COMMENT 'Solo si hay Ariconta.',
  `codmaccl` varchar(10) default NULL,
  `codigiva` smallint(2) NOT NULL default '0',
  `preventa` decimal(10,3) NOT NULL default '0.000',
  `bonigral` decimal(5,4) NOT NULL default '0.0000',
  `canstock` decimal(10,3) NOT NULL default '0.000',
  `stockinv` decimal(10,3) default '0.000',
  `fechainv` date default NULL,
  `preciopmp` decimal(10,3) default NULL,
  `ultpreci` decimal(10,3) default NULL,
  `ultfecha` date default NULL,
  `impuesto` decimal(5,4) default NULL COMMENT 'Solo Familia Carburantes.',
  `tipogaso` tinyint(1) NOT NULL default '0' COMMENT '0=NO 1=Gasolinas 2=Gas.A 3=Gas.ByC',
  PRIMARY KEY  (`codartic`),
  KEY `codfamia` (`codfamia`),
  CONSTRAINT `sartic_fk` FOREIGN KEY (`codfamia`) REFERENCES `sfamia` (`codfamia`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

/*Data for the table `sartic` */


/*Table structure for table `sbanco` */

CREATE TABLE `sbanco` (
  `codbanpr` smallint(2) NOT NULL default '0',
  `nombanco` varchar(35) NOT NULL default '',
  `dombanco` varchar(35) default NULL,
  `codposta` varchar(6) default NULL,
  `pobbanco` varchar(30) default NULL,
  `probanco` varchar(30) default NULL,
  `perbanco` varchar(35) default NULL,
  `telbanco` varchar(10) default NULL,
  `faxbanco` varchar(10) default NULL,
  `wwwbanco` varchar(40) default NULL,
  `maibanco` varchar(40) default NULL,
  `codbanco` varchar(4) NOT NULL default '',
  `codsucur` varchar(4) NOT NULL default '',
  `digcontr` char(2) NOT NULL default '',
  `cuentaba` varchar(10) NOT NULL default '',
  `sufijoem` char(2) default NULL,
  `codmacta` varchar(10) default NULL COMMENT 'Solo si hay Ariconta.',
  PRIMARY KEY  (`codbanpr`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

/*Data for the table `sbanco` */


/*Table structure for table `sbonif` */

CREATE TABLE `sbonif` (
  `codartic` int(6) NOT NULL default '0',
  `tipsocio` tinyint(1) NOT NULL default '0' COMMENT '0=Particular 1=Profesional 2=Comercio',
  `numlinea` smallint(2) NOT NULL default '0',
  `desdecan` int(10) NOT NULL default '0',
  `hastacan` int(10) NOT NULL default '0',
  `bonifica` decimal(10,3) NOT NULL default '0.000',
  PRIMARY KEY  (`codartic`,`tipsocio`,`numlinea`),
  KEY `codartic` (`codartic`),
  CONSTRAINT `sbonif_fk` FOREIGN KEY (`codartic`) REFERENCES `sartic` (`codartic`) ON DELETE CASCADE ON UPDATE CASCADE
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

/*Data for the table `sbonif` */

/*Table structure for table `scaalb` */

CREATE TABLE `scaalb` (
  `codclave` int(6) NOT NULL default '0',
  `codsocio` int(6) NOT NULL default '0',
  `numtarje` int(8) NOT NULL default '0',
  `numalbar` varchar(8) default NULL,
  `fecalbar` date NOT NULL default '0000-00-00',
  `horalbar` time NOT NULL default '00:00:00',
  `codturno` tinyint(1) NOT NULL default '0',
  `codartic` int(6) NOT NULL default '0',
  `cantidad` decimal(8,3) NOT NULL default '0.000',
  `preciove` decimal(8,3) NOT NULL default '0.000',
  `importel` decimal(10,2) NOT NULL default '0.00',
  `codforpa` smallint(2) NOT NULL default '0',
  `matricul` varchar(10) default NULL,
  `codtraba` smallint(4) NOT NULL default '0',
  `numfactu` int(7) NOT NULL default '0',
  `numlinea` smallint(3) NOT NULL default '0',
  PRIMARY KEY  (`codclave`),
  KEY `codsocio` (`codsocio`),
  KEY `codartic` (`codartic`),
  KEY `codforpa` (`codforpa`),
  KEY `codtraba` (`codtraba`),
  CONSTRAINT `scaalb_fk` FOREIGN KEY (`codsocio`) REFERENCES `ssocio` (`codsocio`),
  CONSTRAINT `scaalb_fk1` FOREIGN KEY (`codartic`) REFERENCES `sartic` (`codartic`),
  CONSTRAINT `scaalb_fk2` FOREIGN KEY (`codforpa`) REFERENCES `sforpa` (`codforpa`),
  CONSTRAINT `scaalb_fk3` FOREIGN KEY (`codtraba`) REFERENCES `straba` (`codtraba`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

/*Data for the table `scaalb` */

/*Table structure for table `schfac` */

CREATE TABLE `schfac` (
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

/*Data for the table `schfac` */


/*Table structure for table `scoope` */

CREATE TABLE `scoope` (
  `codcoope` smallint(3) NOT NULL default '0',
  `nomcoope` varchar(40) NOT NULL default '',
  `domcoope` varchar(40) default NULL,
  `codposta` varchar(6) default NULL,
  `pobcoope` varchar(35) default NULL,
  `procoope` varchar(30) default NULL,
  `nifcoope` varchar(9) default NULL,
  `telcoope` varchar(10) default NULL,
  `faxcoope` varchar(10) default NULL,
  `maicoope` varchar(40) default NULL,
  `tipfactu` tinyint(1) unsigned NOT NULL default '0' COMMENT '0=Tarj. 1=Cli 2=Res. 3=Res.Otros.',
  `tipconta` tinyint(1) unsigned NOT NULL default '0' COMMENT '0=Cta.Soc 1=Cta.Cli',
  PRIMARY KEY  (`codcoope`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1 COMMENT='Coop./Colectivos';

/*Data for the table `scoope` */


/*Table structure for table `scryst` */

CREATE TABLE `scryst` (
  `codcryst` smallint(4) unsigned NOT NULL default '0',
  `nomcryst` varchar(30) NOT NULL default '',
  `documrpt` varchar(100) NOT NULL default '',
  `codigiso` varchar(10) default NULL,
  `codigrev` tinyint(2) unsigned default NULL,
  `lineapi1` varchar(140) default NULL,
  `lineapi2` varchar(140) default NULL,
  `lineapi3` varchar(140) default NULL,
  `lineapi4` varchar(140) default NULL,
  `lineapi5` varchar(140) default NULL,
  PRIMARY KEY  (`codcryst`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1 COMMENT='Tipos de Documentos';

/*Data for the table `scryst` */

insert into `scryst` (`codcryst`,`nomcryst`,`documrpt`,`codigiso`,`codigrev`,`lineapi1`,`lineapi2`,`lineapi3`,`lineapi4`,`lineapi5`) values (1,'Factura de clientes','rFactgas.rpt','11112',1,'','','','','');

/*Table structure for table `sempre` */

CREATE TABLE `sempre` (
  `codempre` smallint(3) NOT NULL default '0',
  `nomempre` varchar(40) NOT NULL default '',
  `domempre` varchar(40) NOT NULL default '',
  `codposta` varchar(6) NOT NULL default '',
  `pobempre` varchar(35) NOT NULL default '',
  `proempre` varchar(35) NOT NULL default '',
  `nifempre` varchar(9) NOT NULL default '',
  `telempre` varchar(10) default NULL,
  `faxempre` varchar(10) default NULL,
  `wwwempre` varchar(100) default NULL,
  `maiempre` varchar(100) default NULL,
  `perempre` varchar(40) default NULL,
  PRIMARY KEY  (`codempre`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

/*Data for the table `sempre` */

insert into `sempre` (`codempre`,`nomempre`,`domempre`,`codposta`,`pobempre`,`proempre`,`nifempre`,`telempre`,`faxempre`,`wwwempre`,`maiempre`,`perempre`) values (1,'Alzicoop Carburants S.L.','Plaza Mayor, 30','46600','ALZIRA','VALENCIA','F46024196','9635252522','369696969','http://www.hotfrog.es/Empresas/Coop-Hortofruticola-de-Alzira-Alzicoop',NULL,NULL);

/*Table structure for table `sfamia` */

CREATE TABLE `sfamia` (
  `codfamia` smallint(3) NOT NULL default '0',
  `nomfamia` varchar(25) NOT NULL default '',
  `tipfamia` tinyint(1) NOT NULL default '0' COMMENT '0=Normal 1=Comb. 2=Dto. Solo 1 de tipo 1 y 2.',
  PRIMARY KEY  (`codfamia`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

/*Data for the table `sfamia` */

/*Table structure for table `sforpa` */

CREATE TABLE `sforpa` (
  `codforpa` smallint(2) NOT NULL default '0',
  `nomforpa` varchar(30) NOT NULL default '',
  `tipforpa` tinyint(1) NOT NULL default '0' COMMENT '0=Efec 1=Tr 2=Ta 3=Pa 4=R.B 5=Conf.',
  `cuadresn` tinyint(1) NOT NULL default '0',
  `codmacta` varchar(10) default NULL COMMENT 'Solo si cuadre=1 y hay Ariconta.',
  `contabilizasn` tinyint(1) unsigned NOT NULL default '0',
  PRIMARY KEY  (`codforpa`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1 ROW_FORMAT=DYNAMIC;

/*Data for the table `sforpa` */


/*Table structure for table `slhfac` */

CREATE TABLE `slhfac` (
  `letraser` char(1) NOT NULL default '',
  `numfactu` int(7) unsigned NOT NULL default '0',
  `fecfactu` date NOT NULL default '0000-00-00',
  `numlinea` smallint(4) NOT NULL default '0',
  `numalbar` varchar(8) NOT NULL default '',
  `fecalbar` date NOT NULL default '0000-00-00',
  `horalbar` time NOT NULL default '00:00:00',
  `codturno` smallint(1) NOT NULL default '0',
  `numtarje` int(8) NOT NULL default '0',
  `codartic` int(6) NOT NULL default '0',
  `cantidad` decimal(10,2) NOT NULL default '0.00',
  `preciove` decimal(10,3) NOT NULL default '0.000',
  `implinea` decimal(10,2) NOT NULL default '0.00',
  PRIMARY KEY  (`letraser`,`numfactu`,`fecfactu`,`numlinea`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

/*Data for the table `slhfac` */

/*Table structure for table `smatri` */

CREATE TABLE `smatri` (
  `codsocio` int(6) NOT NULL default '0',
  `numlinea` smallint(2) NOT NULL default '0',
  `matricul` varchar(10) NOT NULL default '',
  `observac` varchar(30) default NULL,
  PRIMARY KEY  (`codsocio`,`numlinea`),
  KEY `codsocio` (`codsocio`),
  CONSTRAINT `smatri_fk` FOREIGN KEY (`codsocio`) REFERENCES `ssocio` (`codsocio`) ON DELETE CASCADE ON UPDATE CASCADE
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

/*Data for the table `smatri` */

/*Table structure for table `sparam` */

CREATE TABLE `sparam` (
  `codparam` smallint(1) NOT NULL default '0',
  `serconta` varchar(20) default NULL,
  `usuconta` varchar(20) default NULL,
  `pasconta` varchar(20) default NULL,
  `numconta` smallint(2) default NULL,
  `ctaconta` varchar(10) default NULL COMMENT 'Cta.Contable contado',
  `ctanegtat` varchar(10) default NULL COMMENT 'Cta.Dif.Negativas.',
  `ctaposit` varchar(10) default NULL COMMENT 'Cta:dif.Positivas.',
  `ctaimpue` varchar(10) default NULL COMMENT 'Cta.Contable impuesto',
  `teximpue` varchar(100) default NULL,
  `bonifact` tinyint(1) NOT NULL default '0' COMMENT '0=NO 1=SI',
  `articdto` int(6) default NULL COMMENT 'Codigo de articulo de descuento',
  `raizctasoc` varchar(10) default NULL,
  `raizctacli` varchar(10) default NULL,
  `ctafamdefecto` varchar(10) default NULL,
  `websoporte` varchar(100) default NULL,
  PRIMARY KEY  (`codparam`),
  KEY `articdto` (`articdto`),
  CONSTRAINT `sparam_fk` FOREIGN KEY (`articdto`) REFERENCES `sartic` (`codartic`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

/*Data for the table `sparam` */

insert into `sparam` (`codparam`,`serconta`,`usuconta`,`pasconta`,`numconta`,`ctaconta`,`ctanegtat`,`ctaposit`,`ctaimpue`,`teximpue`,`bonifact`,`articdto`,`raizctasoc`,`raizctacli`,`ctafamdefecto`,`websoporte`) values (1,'webserver','root','aritel',7,'224000000','226000000','119100000','201000000','hola',0,NULL,'1320','1000','201000000','www.ariadnasoftware.com');

/*Table structure for table `srecau` */

CREATE TABLE `srecau` (
  `fechatur` date NOT NULL default '0000-00-00',
  `codturno` smallint(1) NOT NULL default '0',
  `codforpa` smallint(2) NOT NULL default '0',
  `importel` decimal(10,2) NOT NULL default '0.00',
  `intconta` tinyint(1) NOT NULL default '0',
  PRIMARY KEY  (`fechatur`,`codturno`,`codforpa`),
  KEY `codforpa` (`codforpa`),
  CONSTRAINT `srecau_fk` FOREIGN KEY (`codforpa`) REFERENCES `sforpa` (`codforpa`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

/*Data for the table `srecau` */

/*Table structure for table `ssitua` */

CREATE TABLE `ssitua` (
  `codsitua` smallint(2) NOT NULL default '0',
  `nomsitua` varchar(30) NOT NULL default '',
  `tipsitua` tinyint(1) NOT NULL default '0' COMMENT '0=NO Bloquea 1=Bloquea cliente',
  PRIMARY KEY  (`codsitua`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1 COMMENT='Situaciones clientes';

/*Data for the table `ssitua` */

insert into `ssitua` (`codsitua`,`nomsitua`,`tipsitua`) values (0,'ACTIVO',0);
insert into `ssitua` (`codsitua`,`nomsitua`,`tipsitua`) values (1,'ACTIVO',0);
insert into `ssitua` (`codsitua`,`nomsitua`,`tipsitua`) values (2,'MOROSO',0);

/*Table structure for table `ssocio` */

CREATE TABLE `ssocio` (
  `codsocio` int(6) NOT NULL default '0',
  `codcoope` smallint(3) NOT NULL default '0',
  `nomsocio` varchar(40) NOT NULL default '',
  `domsocio` varchar(40) NOT NULL default '',
  `codposta` varchar(6) NOT NULL default '',
  `pobsocio` varchar(35) NOT NULL default '',
  `prosocio` varchar(35) NOT NULL default '',
  `nifsocio` varchar(9) NOT NULL default '',
  `telsocio` varchar(10) default NULL,
  `faxsocio` varchar(10) default NULL,
  `movsocio` varchar(10) default NULL,
  `maisocio` varchar(40) default NULL,
  `wwwsocio` varchar(40) default NULL,
  `fechaalt` date NOT NULL default '0000-00-00',
  `fechabaj` date default NULL,
  `codtarif` smallint(1) NOT NULL default '0',
  `codbanco` varchar(4) default NULL,
  `codsucur` varchar(4) default NULL,
  `digcontr` char(2) default NULL,
  `cuentaba` varchar(10) default NULL,
  `impfactu` tinyint(1) NOT NULL default '0' COMMENT '0=No imprime 1=Si imprime',
  `dtolitro` decimal(5,4) NOT NULL default '0.0000',
  `codforpa` smallint(2) NOT NULL default '0',
  `tipsocio` tinyint(1) unsigned NOT NULL default '0' COMMENT '0=Particular 1=Profesional 2=Comercio',
  `bonifbas` tinyint(1) NOT NULL default '0' COMMENT '0=NO 1=SI',
  `bonifesp` tinyint(1) NOT NULL default '0' COMMENT '0=NO 1=SI',
  `codsitua` smallint(2) NOT NULL default '0',
  `codmacta` varchar(10) default NULL COMMENT 'Solo si hay Ariconta.',
  `obssocio` varchar(250) default NULL,
  PRIMARY KEY  (`codsocio`),
  KEY `codforpa` (`codforpa`),
  KEY `codsitua` (`codsitua`),
  KEY `codcoope` (`codcoope`),
  KEY `tipsocio` (`tipsocio`),
  CONSTRAINT `ssocio_fk` FOREIGN KEY (`codforpa`) REFERENCES `sforpa` (`codforpa`),
  CONSTRAINT `ssocio_fk1` FOREIGN KEY (`codsitua`) REFERENCES `ssitua` (`codsitua`),
  CONSTRAINT `ssocio_fk2` FOREIGN KEY (`codcoope`) REFERENCES `scoope` (`codcoope`),
  CONSTRAINT `ssocio_fk3` FOREIGN KEY (`tipsocio`) REFERENCES `tiposoci` (`tiposoci`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1 COMMENT='Socios - Clientes';

/*Data for the table `ssocio` */

/*Table structure for table `starif` */

CREATE TABLE `starif` (
  `codartic` int(6) NOT NULL default '0',
  `codtarif` smallint(1) NOT NULL default '0',
  `preventa` decimal(10,3) NOT NULL default '0.000',
  PRIMARY KEY  (`codartic`,`codtarif`),
  KEY `codartic` (`codartic`),
  CONSTRAINT `starif_fk` FOREIGN KEY (`codartic`) REFERENCES `sartic` (`codartic`) ON DELETE CASCADE ON UPDATE CASCADE
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

/*Data for the table `starif` */

/*Table structure for table `starje` */

CREATE TABLE `starje` (
  `codsocio` int(6) NOT NULL default '0',
  `numlinea` smallint(2) unsigned NOT NULL default '0',
  `numtarje` int(8) NOT NULL default '0',
  `nomtarje` varchar(40) NOT NULL default '',
  `codbanco` varchar(4) default NULL,
  `codsucur` varchar(4) default NULL,
  `digcontr` char(2) default NULL,
  `cuentaba` varchar(10) default NULL,
  `tiptarje` tinyint(1) NOT NULL default '0' COMMENT '0=Normal 1=Bonificado',
  PRIMARY KEY  (`codsocio`,`numlinea`),
  UNIQUE KEY `numtarje` (`numtarje`),
  KEY `codsocio` (`codsocio`),
  CONSTRAINT `starje_fk` FOREIGN KEY (`codsocio`) REFERENCES `ssocio` (`codsocio`) ON DELETE CASCADE ON UPDATE CASCADE
) ENGINE=InnoDB DEFAULT CHARSET=latin1 COMMENT='Tarjetas clientes';

/*Data for the table `starje` */

/*Table structure for table `stipom` */

CREATE TABLE `stipom` (
  `codtipom` char(3) NOT NULL default '0',
  `nomtipom` varchar(30) NOT NULL default '',
  `contador` mediumint(7) NOT NULL default '0',
  `letraser` char(1) NOT NULL default '',
  PRIMARY KEY  (`codtipom`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1 COMMENT='Tipos de Movimiento';

/*Data for the table `stipom` */

insert into `stipom` (`codtipom`,`nomtipom`,`contador`,`letraser`) values ('FAB','FACTURA GASOLEO BONIFICADO',0,'B');
insert into `stipom` (`codtipom`,`nomtipom`,`contador`,`letraser`) values ('FAG','FACTURA GASOLINA',161,'A');
insert into `stipom` (`codtipom`,`nomtipom`,`contador`,`letraser`) values ('FAT','FACTURA TPV',0,'F');

/*Table structure for table `straba` */

CREATE TABLE `straba` (
  `codtraba` smallint(4) NOT NULL default '0',
  `nomtraba` varchar(30) NOT NULL default '',
  `domtraba` varchar(30) default NULL,
  `codpobla` varchar(6) default NULL,
  `pobtraba` varchar(30) default NULL,
  `protraba` varchar(30) default NULL,
  `niftraba` varchar(9) default NULL,
  `teltraba` varchar(10) default NULL,
  `movtraba` varchar(10) default NULL,
  `cartraba` varchar(30) default NULL,
  `maitraba` varchar(40) default NULL,
  `loginweb` varchar(20) NOT NULL default '',
  `passwweb` varchar(20) NOT NULL default '',
  PRIMARY KEY  (`codtraba`),
  KEY `itraba1` (`nomtraba`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1 COMMENT='Trabajadores';

/*Data for the table `straba` */

insert into `straba` (`codtraba`,`nomtraba`,`domtraba`,`codpobla`,`pobtraba`,`protraba`,`niftraba`,`teltraba`,`movtraba`,`cartraba`,`maitraba`,`loginweb`,`passwweb`) values (0,'Ariadna Software','Franco Tormo','46007','Valencia','Valencia','B9699999','aaa',NULL,NULL,NULL,'root','aritel');

/*Table structure for table `sturno` */

CREATE TABLE `sturno` (
  `fechatur` date NOT NULL default '0000-00-00',
  `codturno` smallint(1) NOT NULL default '0',
  `numlinea` smallint(4) NOT NULL default '0',
  `tiporegi` tinyint(1) NOT NULL default '0' COMMENT '0=Contadores 1=Tanque 2=Ventas tipo 3=Compra 4=Varillas',
  `numtanqu` smallint(3) default NULL,
  `nummangu` smallint(3) default NULL,
  `codartic` int(6) default NULL,
  `tipocred` tinyint(1) default NULL,
  `litrosve` decimal(10,2) default NULL,
  `importel` decimal(10,2) default NULL,
  `containi` decimal(10,2) default NULL,
  `contafin` decimal(10,2) default NULL,
  PRIMARY KEY  (`fechatur`,`codturno`,`numlinea`),
  KEY `fechatur` (`fechatur`,`codturno`),
  KEY `codartic` (`codartic`),
  CONSTRAINT `sturno_fk` FOREIGN KEY (`codartic`) REFERENCES `sartic` (`codartic`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

/*Data for the table `sturno` */


/*Table structure for table `tipocred` */

CREATE TABLE `tipocred` (
  `tipocred` tinyint(1) NOT NULL default '0',
  `nomcredi` char(20) NOT NULL default ''
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

/*Data for the table `tipocred` */

insert into `tipocred` (`tipocred`,`nomcredi`) values (0,'Clientes');

/*Table structure for table `tipofami` */

CREATE TABLE `tipofami` (
  `tipfamia` tinyint(4) NOT NULL default '0',
  `destipfa` varchar(20) NOT NULL default '',
  PRIMARY KEY  (`tipfamia`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

/*Data for the table `tipofami` */

insert into `tipofami` (`tipfamia`,`destipfa`) values (0,'Productos');
insert into `tipofami` (`tipfamia`,`destipfa`) values (1,'Combustible');
insert into `tipofami` (`tipfamia`,`destipfa`) values (2,'Bonificaciones');

/*Table structure for table `tiporegi` */

CREATE TABLE `tiporegi` (
  `tiporegi` tinyint(1) NOT NULL default '0',
  `nomtipre` char(20) NOT NULL default ''
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

/*Data for the table `tiporegi` */

insert into `tiporegi` (`tiporegi`,`nomtipre`) values (0,'Contadores');
insert into `tiporegi` (`tiporegi`,`nomtipre`) values (1,'Tanques');
insert into `tiporegi` (`tiporegi`,`nomtipre`) values (2,'Ventas tipo');
insert into `tiporegi` (`tiporegi`,`nomtipre`) values (3,'Compras');
insert into `tiporegi` (`tiporegi`,`nomtipre`) values (4,'Varillas');

/*Table structure for table `tiposoci` */

CREATE TABLE `tiposoci` (
  `tiposoci` tinyint(1) unsigned NOT NULL default '0',
  `nomtipso` varchar(20) NOT NULL default '',
  PRIMARY KEY  (`tiposoci`),
  UNIQUE KEY `tiposoci` (`tiposoci`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

/*Data for the table `tiposoci` */

insert into `tiposoci` (`tiposoci`,`nomtipso`) values (0,'Particular');
insert into `tiposoci` (`tiposoci`,`nomtipso`) values (1,'Profesional');
insert into `tiposoci` (`tiposoci`,`nomtipso`) values (2,'Comercio');

/*Table structure for table `tmpinformes` */

CREATE TABLE `tmpinformes` (
  `codusu` smallint(3) unsigned NOT NULL default '0',
  `codigo1` int(6) unsigned NOT NULL default '0',
  `fecha1` date default NULL,
  `fecha2` date default NULL,
  `campo1` smallint(4) unsigned default NULL,
  `campo2` smallint(4) unsigned default NULL,
  `nombre1` varchar(40) default NULL,
  `importe1` decimal(12,2) default NULL,
  `importe2` decimal(12,2) default NULL,
  `importe3` decimal(12,2) default NULL,
  `importe4` decimal(12,2) default NULL,
  `importe5` decimal(12,2) default NULL,
  `porcen1` decimal(5,2) default NULL,
  `porcen2` decimal(5,2) default NULL,
  `importeb1` decimal(12,2) default NULL,
  `importeb2` decimal(12,2) default NULL,
  `importeb3` decimal(12,2) default NULL,
  `importeb4` decimal(12,2) default NULL,
  `importeb5` decimal(12,2) default NULL
) ENGINE=MyISAM DEFAULT CHARSET=latin1 COMMENT='Temporal para informes';

/*Data for the table `tmpinformes` */

/*Table structure for table `usuarios` */

CREATE TABLE `usuarios` (
  `codusu` smallint(1) unsigned NOT NULL default '0',
  `nomusu` char(30) NOT NULL default '',
  `dirfich` char(255) default NULL,
  `nivelusu` tinyint(1) NOT NULL default '-1',
  `login` char(20) NOT NULL default '',
  `passwordpropio` char(20) NOT NULL default '',
  `nivelusuges` tinyint(4) NOT NULL default '-1',
  PRIMARY KEY  (`codusu`)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;

/*Data for the table `usuarios` */

insert into `usuarios` (`codusu`,`nomusu`,`dirfich`,`nivelusu`,`login`,`passwordpropio`,`nivelusuges`) values (0,'root',NULL,0,'root','aritel',0);

/*Table structure for table `zbloqueos` */

CREATE TABLE `zbloqueos` (
  `codusu` smallint(1) unsigned NOT NULL default '0',
  `tabla` char(20) NOT NULL default '',
  `clave` char(30) NOT NULL default '',
  PRIMARY KEY  (`tabla`,`clave`)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;

/*Data for the table `zbloqueos` */

SET SQL_MODE=@OLD_SQL_MODE;
SET FOREIGN_KEY_CHECKS=@OLD_FOREIGN_KEY_CHECKS;
