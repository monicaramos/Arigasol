/* campo para sacar el listado de estadistica de artículos para el Regaixo */

USE `arigasol`;

alter table `ssocio` add column `grupoestartic` tinyint (1)  DEFAULT '0' NOT NULL  COMMENT '0=coop 1=tarj.visa 2=cred.local 3=clientes paso 4=efectivo' after `obssocio`;

/*ampliamos la tabla auxiliar par aobtener el informe*/

alter table `tmpinformes` add column `importe6` decimal (12,2)   NULL  after `importe5`, add column `importeb6` decimal (12,2)   NULL  after `importeb5`