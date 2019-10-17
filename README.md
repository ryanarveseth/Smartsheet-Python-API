# Smartsheet-Python-API
How to connect to your Smartsheet with a Python API. 

### Add Rows and Attachments

This file helps the user add rows and attachments to a Smartsheet (almost) at the same time. 
When running this file on the command line or terminal, it takes 6 parameters:
- email    ---> contact column
- rType    ---> dropdown (request type)
- pat      ---> optional field, default: 0
- priority ---> (1 - 3)
- attachment ---> (file name)
- att_path ---> (file path)

#### Functions
```python
def get_cell_by_column_name(row, column_name):
```
Returns the cell id for a given row and column (by column *name*)

```python
def initial(i_email, i_rType, i_pat, i_priority, i_attachment, i_att_path):
```
Sets all of the global variables in the program. When the program starts, it calls this function if there are the correct number of argv's. If not, the program will stop. In this case, we'll always expect 7 arguments.

```python
def addRow(rType, pat, priority, email):
```
addRow takes 4 parameters, and adds each value to the corresponding column in the smartsheet

```python 
def getAddedRow(email, rType, pat, priority):
```
getAddedRow returns the Row ID of the highest row number that matches all of your parameters you passed with addRow();
IF no row matches our criteria, it will return "None".
