alter table `arigasol`.`scaalb` change `horalbar` `horalbar` datetime  DEFAULT '00:00:00' NOT NULL ;
alter table `arigasol`.`slhfac` change `horalbar` `horalbar` datetime  DEFAULT '00:00:00' NOT NULL ;
update scaalb set horalbar = fecalbar;
update slhfac set horalbar = fecalbar;