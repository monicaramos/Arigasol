USE `arigasol`;


/* campos necesarios para declaracion de gasoleo profesional*/
alter table `sparam` add column `cim` varchar (8)   NULL  COMMENT 'Codigo gasolinera GP' after `smtppass`, 
                     add column `cee` varchar (4)   NULL  COMMENT 'GP' after `cim`;


/* campos de articulos de combustible para gasoleo profesional */
alter table `sartic` add column `gp` tinyint (1)  DEFAULT '0' NOT NULL  COMMENT '0=no se declara 1=se declara como GP' after `tipogaso`, 
                     add column `porcbd` decimal (6,2)  DEFAULT '0' NOT NULL  COMMENT 'porcentaje de biodiesel' after `gp`;


/* campo de la scaalb para saber si está declarado o no */
alter table `scaalb` add column `declaradogp` tinyint (1)  DEFAULT '0' NOT NULL  COMMENT '0=NO 1=SI' after `numlinea`;




