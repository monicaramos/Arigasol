/* modificacion de la Tabla temporal  */
use `arigasol`;

alter table `tmpinformes` add column `precio1` decimal (12,3)   NULL  after `nombre2`;
