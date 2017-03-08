# robotframework-excelreader
a python excel reader for [robotframework](http://robotframework.org/)

## Installation

```
$ pip install robotframework
$ pip install openpyxl
````

## Usage
Refer to demo.txt

## Documentation

**read_multiple_records**

*[Arguments]*  path, keyDict, hasHeader=True, debug=False

*[Documentation]*  Reads multiple records in xlsx. Returns a list of dictionary.

**read_single_record**

*[Arguments]*  path, keyDict, row='1', debug=False

*[Documentation]*  Reads the record defined by the row in xlsx. Returns a list of one dictionary item.

**read_single_row**

*[Arguments]*  path, keyDict, row='1', debug=False

([Documentation]*  Reads the record defined by the row in xlsx. Returns a list of attributes.