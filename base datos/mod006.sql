use `arigasol`;

CREATE TABLE `cabrangos` (                                     
             `codigo` smallint(3) NOT NULL default '0' COMMENT 'codigo',  
             `descripcion` varchar(50) NOT NULL,                          
             PRIMARY KEY  (`codigo`)                                      
           ) ENGINE=InnoDB DEFAULT CHARSET=latin1 ;
                        
CREATE TABLE `linrangos` (                      
             `codigo` smallint(3) NOT NULL default '0',    
             `numlinea` smallint(3) NOT NULL default '0',  
             `desdehora` datetime NOT NULL,                
             `hastahora` datetime NOT NULL,                
             PRIMARY KEY  (`codigo`,`numlinea`)            
           ) ENGINE=InnoDB DEFAULT CHARSET=latin1;          