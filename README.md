Convert xls files into sql script

## Dependency

* python 2.x
* python lib: xlrd

## Usage

command(python generateSQL.py <xls file name> [output file name，default：result.sql])

## How it works

It will read excel sheets one by one. Each sheet as a database table. Read the first row of each sheet as the table name, will stop when first meet a blank cell in this row. Then read rest rows as data, empty row will be skiped.
