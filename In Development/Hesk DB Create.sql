

CREATE DATABASE `hesk` 
DEFAULT CHARACTER SET utf8 COLLATE utf8_general_ci;

CREATE USER 'hesk_admin'@'localhost'; 
CREATE USER 'hesk_admin'@'127.0.0.1';
CREATE USER 'hesk_admin'@'::1';

SET PASSWORD 
FOR 'hesk_admin'@'localhost' = PASSWORD('rUq6GwZ23c!HkkmbKWdn');
SET PASSWORD 
FOR 'hesk_admin'@'127.0.0.1' = PASSWORD('rUq6GwZ23c!HkkmbKWdn');
SET PASSWORD 
FOR 'hesk_admin'@'::1' = PASSWORD('rUq6GwZ23c!HkkmbKWdn');


GRANT ALL PRIVILEGES ON 
`hesk`.* TO 'hesk_admin'@'localhost' WITH GRANT OPTION;
GRANT ALL PRIVILEGES ON 
`hesk`.* TO 'hesk_admin'@'127.0.0.1' WITH GRANT OPTION;
GRANT ALL PRIVILEGES ON 
`hesk`.* TO 'hesk_admin'@'::1' WITH GRANT OPTION;

