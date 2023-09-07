""" 
This is the main file for the project. It will be used to run the program.
First, we need to import the 3 dependancies
1. Pandas
2. Requests
3. openpyxl 
4. requests_toolbelt
You can do this by running the following commands in your terminal:
pip install pandas && pip install requests && pip install openpyxl
or 
pip install pandas
pip install requests
pip install openpyxl
"""
# import xlsxwriter module
import xlsxwriter
import requests
import pandas as pd
from requests_toolbelt.utils import dump


filename = 'origin.xlsx'

df = pd.read_excel(filename, sheet_name='Sheet2')
# filter_df = pd.isnull(df.var2)
print(df)

entries = []
for rowIndex, row in df.iterrows(): #iterate over rows
    for columnIndex, value in row.items():
        # print(value, end="\t")
        # print(value, end="\n")
        val = str(value)
        if val != 'nan':
            entries.append(val)
    print()


print(entries)

url = 'https://openscoring.du.edu/'
model = '/llm?model=gpt-davinci-paper_alpha'
prompt = '&prompt=brick'
inputs = [input.replace(' ', '%20') for input in entries]

print(inputs)

# response = requests.get('&input=build%20a%20wall&input=paper%20weight&input=weapon&input_type=csv')

# Workbook() takes one, non-optional, argument
# which is the filename that we want to create.
workbook = xlsxwriter.Workbook('hello.xlsx')
 
# The workbook object is then used to add new
# worksheet via the add_worksheet() method.
worksheet = workbook.add_worksheet()

data = response.json()
for i, key in enumerate(data):
    # doubler
    i = 4 * i
    print(key, '->', data[key])
    # worksheet.write(i, key, data[key])
    worksheet.write(0, i, str(key))
    try:
        worksheet.write(0, i, str(key))
        if type(data[key]) == list:
            for j, value in enumerate(data[key]):
                worksheet.write(j+1, i, str(value))
        elif type(data[key]) == dict:
            for j, value in enumerate(data[key]):
                worksheet.write(j+1, i, str(value))
        else:
            worksheet.write(1, i, str(data[key]))
    except:
        print('AAAAAAA', data[key])




 

 
# Use the worksheet object to write
# data via the write() method.

 
# Finally, close the Excel file
# via the close() method.
workbook.close()

# data = dump.dump_all(resp)
# print(data)
# print(data.decode('utf-8'))
# data = data.decode('utf-8')
# print(data.parameters)

# url = 'https://www.googleapis.com/qpxExpress/v1/trips/search?key=mykeyhere'
# payload = open("request.json")
# headers = {'content-type': 'application/json', 'Accept-Charset': 'UTF-8'}
# r = requests.post(url, data=payload, headers=headers)