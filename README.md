#xlsx2json-py
##Usage
export from excel/ to json/  
run it:
```
$ python xlsx2json.py
```
.xlsx file head row default 2 and you can define it by yourself  
don't forget add the num when run it
```
$ python xlsx2json.py 3
```
##Supported data type
* int
* float
* string
* boolean
* date
* array
* object

##Data rule
* use half-angle
* if only one data in array, ending with _,_
* you must set the cell value as a date type when you want to use the date type
