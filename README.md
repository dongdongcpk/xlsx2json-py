#xlsx2json-py
##Env
python version >= 2.7
##Dependencies
openpyxl (2.2.2)  
you can install it by pip, like this: pip install openpyxl
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
* if only one data in array, ending with `,`

##Example
![example](https://github.com/dongdongcpk/xlsx2json-py/blob/master/example/example.png)
result:
```json
[
    {
        "person": {
            "age": 18, 
            "name": "dong"
        }, 
        "value": 0.12, 
        "start": "Mon May 11 00:00:00 2015", 
        "flag": true, 
        "words": [
            1, 
            2, 
            3, 
            4
        ], 
        "id": 1, 
        "desc": "hello"
    }, 
    {
        "person": {
            "id": 5
        }, 
        "value": 2.4, 
        "start": "Mon May 11 13:30:24 2015", 
        "flag": false, 
        "words": [
            1
        ], 
        "id": 2, 
        "desc": "你好"
    }, 
    {
        "person": {
            "max": true, 
            "min": false
        }, 
        "value": 3.14, 
        "start": "Mon Jan 20 00:00:00 2020", 
        "flag": true, 
        "words": [
            "hello", 
            "world"
        ], 
        "id": 3, 
        "desc": "foo"
    }, 
    {
        "person": {
            "level": 100
        }, 
        "value": 2.22, 
        "start": "Fri May  1 02:30:00 2020", 
        "flag": true, 
        "words": [
            true, 
            false
        ], 
        "id": 4, 
        "desc": "world"
    }
]
```