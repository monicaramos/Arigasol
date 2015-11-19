use `arigasol`;

alter table `arigasol`.`sparam` add column `cooperativa` smallint (2)  DEFAULT '1' NOT NULL  COMMENT '1=Alzira 2=Catadau' after `websoporte`