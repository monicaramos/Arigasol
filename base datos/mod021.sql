USE `arigasol`;

ALTER TABLE `sparam` add column `concedebe` smallint(4) NOT NULL DEFAULT 0 after `cee`;
ALTER TABLE `sparam` add column `concehaber` smallint(4) NOT NULL DEFAULT 0 after `concedebe`;
ALTER TABLE `sparam` add column `numdiari` smallint(4) NOT NULL DEFAULT 0 after `concehaber`;
