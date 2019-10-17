import smartsheet.attachments
import smartsheet.home
import smartsheet.cells
import smartsheet.sheets
import logging
import os.path
import datetime
import json
import requests
import glob
import os
import sys

### This is your API Key... this program won't work unless you have a key here
access_token = ""
### Sheet ID
sheet_ID = ""

column_map = {}

# These are the required fields for this Smartsheet case
email = ""
rType = ""
pat = ""
priority = ""
attachment = ""
att_path = ""

# Simple constructor function
def initial(e, t, p, pr, a, ap):
    # Set the global variables
    global email
    email = e
    global rType
    rType = t
    global pat
    pat = p
    global priority
    priority = pr
    global attachment
    attachment = a
    global att_path
    att_path = ap


# Helper function to find cell in a row
def get_cell_by_column_name(row, column_name):
    # returns the cell id for a given row and column
    column_id = column_map[column_name]
    return row.get_column(column_id)


# Takes the request information and creates a new row in our smartsheet.
def addRow(rType, pat, priority, email):
    row_a = smartsheet.models.Row()
    # Add the row at the top
    row_a.to_top = True

    # Here are the cell values
    row_a.cells.append({ 
        'column_id': column_map["Reason Code"], 
        'value': rType,
        'strict': False
    })
    row_a.cells.append({
        'column_id': column_map["PAT #"], 
        'value': pat,
        'strict': False
    })
    row_a.cells.append({ 
        'column_id': column_map["Priority"], 
        'value': priority,
        'strict': False
    })
    row_a.cells.append({ 
        'column_id': column_map["Name of Pricing Person"], 
        'value': email,
        'strict': False
    })
    row_a.cells.append({ 
        'column_id': column_map["Load Status"], 
        'value': 'Submitted',
        'strict': False
    })
    # This line adds the row to the smartsheet
    response = smart.Sheets.add_rows(sheet_ID, row_a)
    print("Row Inserted")


# Return a new Row with updated cell values, else None to leave unchanged
def evaluate_row_and_build_updates(email, rType, pat, priority):
    # Row ID is our unique auto-incrementing primary column. 
    r_id = 0
    # new row ID
    new_row_id = None
    # Loop through all of the rows
    for row in sheet.rows:
        # Find the cells and values we want to evaluate
        status_cell = get_cell_by_column_name(row, "Load Status")
        status_value = status_cell.display_value

        type_cell = get_cell_by_column_name(row, "Reason Code")
        type_value = type_cell.display_value

        email_cell = get_cell_by_column_name(row, "Name of Pricing Person")
        email_value = email_cell.value

        pat_cell = get_cell_by_column_name(row, "PAT #")
        pat_value = pat_cell.display_value

        pri_cell = get_cell_by_column_name(row, "Priority")
        pri_value = pri_cell.display_value

        r_id_cell = get_cell_by_column_name(row, "Row ID")

        #if r_id_cell.display_value != "" and r_id_cell.display_value != None:
        r_id_val = r_id_cell.display_value

        # this is where we check each row for the values we just added.
        # if all of these match, it could be the row that we add our 
        # attachment to
        if str(status_value) == "Submitted" and str(type_value) == str(rType) \
            and str(email_value) == str(email) and str(pat_value) == str(pat) \
                and str(pri_value) == str(priority):
            # change if the row id is greater than previously largest one
            if int(r_id_val) > int(r_id):
                new_row_id = row.id
                r_id = int(r_id_val)
        # return the row to add our attachment to
        return new_row_id


###################################################
# Here is where we start the program.
# When executing the program, we pass 6 arguments. 
#   sys.argv[1] = email
#   sys.argv[2] = request type
#   sys.argv[3] = pricing request number (PAT)
#   sys.argv[4] = priority (1-3)
#   sys.argv[5] = fileName
#   sys.argv[6] = filePath
#
#####################################
print("Starting ...")
if (len(sys.argv) == 7):
    initial(sys.argv[1],sys.argv[2],sys.argv[3],sys.argv[4],sys.argv[5],sys.argv[6])
else:
    sys.exit(0)
    
# Initialize client
smart = smartsheet.Smartsheet(access_token)

# Make sure we don't miss any error
smart.errors_as_exceptions(True)

# Log all calls
logging.basicConfig(filename='rwsheet.log', level=logging.INFO)

# Load entire sheet
sheet = smart.Sheets.get_sheet(sheet_ID)

print("Loaded " + str(len(sheet.rows)) + " rows from sheet: " + sheet.name)

# Build column map for later reference - translates column names to column id
for column in sheet.columns:
    column_map[column.title] = column.id

# Add the row
addRow(rType, pat, priority, email)

# Now that we've added our row, we need to re-load the data
sheet = smart.Sheets.get_sheet(sheet_ID)

# Create our column map again
for column in sheet.columns:
    column_map[column.title] = column.id

# Get the row to update with the attachment
rowToUpdate = evaluate_row_and_build_updates(email, rType, pat, priority)


if rowToUpdate == None:
    print("Something went wrong with uploading the row...")
else:
    print("Adding attachment to sheet id " + str(sheet.id))
    
    # My group only adds excel workbooks. If we didn't, I'd change the 'application/ms-excel' to a variable
    updated_attachment = smart.Attachments.attach_file_to_row(sheet_ID,rowToUpdate,
    (attachment,open(att_path, 'rb'),'application/ms-excel'))
