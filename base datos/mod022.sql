USE `arigasol`;

alter table `ssocio` add column `facturafp` tinyint (1)  DEFAULT '0' NOT NULL  COMMENT '0=NO 1=SI factura con fp de ficha de cliente' after `grupoestartic`;

alter table `sforpa` add column `diasvto` int (4) UNSIGNED  DEFAULT '0' NOT NULL  COMMENT 'Dias a añadir a fecvto' after `permitebonif`;
