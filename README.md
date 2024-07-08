## Merge Excel


Simple project to merge excel workbooks into one large workbook & convert from one excel format to another


### How to install

* make sure you have php installed in your system
* make sure you have [composer](https://getcomposer.org/)
* install project dependecies 

```sh

composer install

```
The program supports two commands `merge` & `convert`

* `merge` - use this to merge multiple workbooks into one

```sh
php index.php merge --path=/path/to/workbooks

```

* `convert` - use this to convert from one Excel format to another

```sh
php index.php convert --workbook=/path/to/workbook --from=xls --to=xlsx

```
* help message, you can simply run the index.php file to get an overview of availble commands

```sh

php index.php # this will dispplay a help guiding you on how to use the program

```

*vualah, that's it, good luck, hope the program helps you


### License
MIT
