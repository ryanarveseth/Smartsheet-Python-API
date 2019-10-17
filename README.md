# Smartsheet-Python-API
How to connect to your Smartsheet with a Python API. 

### Add Rows and Attachments

This file helps the user add rows and attachments (almost) at the same time. 
When running this file on the command line or terminal, it takes 6 parameters:
- email    ---> contact column
- rType    ---> dropdown (request type)
- pat      ---> optional field, default: 0
- priority ---> (1 - 3)
- attachment ---> (file name)
- att_path ---> (file path)

#### Functions



```python
def addRow(rType, pat, priority, email):
```
- addRow takes 4 parameters, and adds each value to the corresponding column in the smartsheet
