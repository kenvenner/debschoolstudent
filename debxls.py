from openpyxl import load_workbook

import sys
import datetime


# set up at the beginning
xls_columns = [
    'Notused',
    'Date',
    'Owed',
    'Paid',
    'Message',
    'Balance',
    'PastDue',
]

# set up at the beginning
xls_columns_out = [
    'filename',
    'worksheet',
    'row',
    'Date',
    'Owed',
    'Paid',
    'Message',
    'Balance',
    'PastDue',
]

# filename
xlsfilename = './Actual Dispersement record.xlsx'

# create array that holds all the read in data
xlsdata = []

# Load in the workbook (set the data_only=True flag to get the value on the formula)
wb = load_workbook(xlsfilename, data_only=True)

# get the list of sheets that need to process
for sheetName in wb.sheetnames:

    # create a workbook sheet object - using the name to get to the right sheet
    s = wb[sheetName]

    # grab the sheet title - not sure i need this - that is already in sheetName
    sheettitle = s.title
    sheetmaxrow = s.max_row
    sheetmaxcol = s.max_column

    # value we are looking for
    cellValue = 'Date'
    
    # the column that has this value
    column = 1
    
    # find the row that has the headers
    for row_header in range(1,6):
        if s.cell(row=row_header, column=column).value == cellValue:
            break
        
    # print out what we found
    #print ('found the matching column:', row_header, ':', column)
        
    # pull in all the data from this sheet that we are interested in 
    for row in range(row_header+1, sheetmaxrow):
        # create a new record to hold this rows data
        rec = {}
        # fill in the major attributes
        rec['worksheet'] = sheetName
        rec['filename'] = xlsfilename
        rec['row'] = str(row)
        # go through the columns of this row
        for col in range(1,7):
            # now populate the record
            rec[xls_columns[col]] = s.cell(row=row, column=col).value
            #print(sheetName,':',row,':',col,':',type(rec[xls_columns[col]]))
            # change the date column to a string
            # if col == 1 and rec[xls_columns[col]] != None and type(rec[xls_columns[col]]) == 'datetime.datetime':
            #if col == 1 and type(rec[xls_columns[col]]) == 'datetime.datetime':
            if isinstance(rec[xls_columns[col]], datetime.datetime):
                #print(sheetName,':',row,':',col,':converted-field')
                rec[xls_columns[col]] = rec[xls_columns[col]].strftime('%m-%d-%Y')
            else:
                rec[xls_columns[col]] = str(rec[xls_columns[col]])
            
        # now show what we got
        #print('rec:',rec)
        
        # now add this record to the current array
        xlsdata.append(rec)

# we are done - print out all the records
#print('all-records:', xlsdata)

# create the headers
#print('--------------------------------------------------------')
print(','.join(xls_columns_out))

# now create an output that is comma delimited
for rec in xlsdata:
    newrec = []
    for colname in xls_columns_out:
        newrec.append( rec[colname] )
    # and then join/output this record
    print(','.join(newrec))

    
sys.exit()

# 
row = 2
# first row of titles
for column in range(1,6):
    print(s.cell(row=row, column=column).value)
