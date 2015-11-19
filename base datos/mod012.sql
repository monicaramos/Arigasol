/*Este cambio se necesitó en Alzira por error en claves referenciales*/
/*para ejecutarlo quitar las almohadillas*/

#USE `conta15`;

/*Cambio de la claves referenciales de scobro y de spago*/

/*scobro*/
#alter table `conta15`.`scobro` drop foreign key `scobro_ibfk_1` 
#alter table `conta15`.`scobro` add foreign key `FK_scobro`(`codforpa`) references `sforpa` (`codforpa`) on delete restrict  on update restrict 

/*spago*/
#alter table `conta15`.`spagop` drop foreign key `spagop_ibfk_1` 
#alter table `conta15`.`spagop` add foreign key `FK_spagop`(`codforpa`) references `sforpa` (`codforpa`) on delete restrict  on update restrict 
