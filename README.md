mysql2excel
=============

Command-line tool to dump MySQL tables to Microsoft Excel .xlsx Spreadsheets written
in Java.


## Downloading

Download the latest release from https://github.com/DevDungeon/mysql2excel/releases
or clone the repo from https://github.com/DevDungeon/mysql2excel
    
    git clone https://github.com/DevDungeon/mysql2excel

    # Compile and package with Maven
    cd mysql2excel
    mvn package

## Usage

### Print help information

    java -jar mysql2excel-1.0.2.jar -h

### Generate a settings template file

    java -jar mysql2excel-1.0.2.jar -g sample.config

### Dump from MySQL to Excel using settings in config file

    java -jar mysql2excel-1.0.2.jar my.config


## Project Page

* https://www.github.com/hezengliang/mysql2excel

## Contact

NanoDano <nanodano@devdungeon.com>  
He Zengliang <hezengliang@hotmail.com>  

## Changelog
* 2020-12-01 v1.0.2
    * enhancement: changed tablename to array
    * enhancement: added condition for all tables
    * enhancement: add replacement with ~date~ and ~time~ for output file name and auto creating directory

* 2018-03-25 v1.0.1
    * Bug fix: added convert zeroDateTimeBehavior to convertToNull, minor formatting tweaks
* 2018-03-18 v1.0.0
    * Initial stable release

## To do

* Allow multiple tables or all tables to be dumped at once
* Allow custom SQL query to be run
