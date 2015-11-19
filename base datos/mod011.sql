
USE `arigasol`;

/*Cambio de la estructura de sparam  */

alter TABLE `sparam` add(
  `diremail` varchar(50) default NULL,
  `smtphost` varchar(50) default NULL,
  `smtpuser` varchar(50) default NULL,
  `smtppass` varchar(50) default NULL
)