/* Tabla temporal para la reimpresion de facturas, por caso de querer reimprimir facturas que tengan 
alg�n art�culo de gasoleo B */

USE `arigasol`;

alter table `arigasol`.`schfac` add column `observac` varchar (72)   NULL  after `intconta`

