use `arigasol`;

alter table `arigasol`.`starje` add column `matricul` varchar (10)   NULL  after `tiptarje`,change `tiptarje` `tiptarje` tinyint (1)  DEFAULT '0' NOT NULL  COMMENT '0=Normal 1=Bonificado 2=Profesional';

alter table `arigasol`.`soltarje` change `tipotarje` `tiptarje` smallint (1)  DEFAULT '0' NOT NULL  COMMENT 'tipo:0=normal 1=bonificada'