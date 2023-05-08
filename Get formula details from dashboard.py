

import json
import re
import pandas as pd
import requests

# Get data using API

user_name = 'usr' #your admin user name. 
password = 'pwd' #your password.
Sisense_url = "https://____.us" #Sisense address.
dashboard_id = 'id'

#------------------------------------------------------------------------------

#Get auth_token

url = Sisense_url+"/api/v1/authentication/login"

payload = "username="+user_name.replace('@','%40')+"&password="+password
headers = {
  'cache-control': "no-cache",
  'content-type': "application/x-www-form-urlencoded"
}

response = requests.request("POST", url, data=payload, headers=headers)
REST_API_TOKEN = response.json()["access_token"]

# #------------------------------------------------------------------------------

url = Sisense_url+f"/api/v1/dashboards/{dashboard_id}/export/dash"
headers = {'Authorization': "Bearer " + REST_API_TOKEN}
response = requests.get(url, headers=headers)
data=response.json()
name = data['title']


#Get data using file instead of API

# name = "escaped"
# filename_in = f"{name}.dash"
# f = open(filename_in,)
# data = json.load(f)


blank_cell = "" 
rows_list = []
widget_title_list = [] #for checking for dupes
dashboard_filter_list = ["Dashboard Filters"]
formula_list = []

#Create dictionaries for mapping the formula alias to the formula details. One dictionary excludes the filters to create the high-level formula summary
def create_formula_dict(data):
    formula_dict = {}
    formula_dict_no_filter = {}
    for widget,widget_value in enumerate(data['widgets']):
        for panel,panel_value in enumerate(widget_value['metadata']['panels']):
            for item,item_value in enumerate(panel_value['items']):
                if 'context' in item_value['jaql']:
                    for context,context_value in item_value['jaql']['context'].items():
                        if 'filter' in context_value:
                            filter_dict = {key:val for key, val in context_value['filter'].items() if (key != 'explicit' and key != 'multiSelection' and key != 'by')}
                            formula_dict[context] = [context_value['dim'],filter_dict] 
                            formula_dict_no_filter[context] = blank_cell
                        elif 'formula' in context_value:
                            formula_dict[context] = context_value['title']
                            formula_dict_no_filter[context] = context_value['title']
                        else:
                            formula_dict[context] = context_value['dim']
                            formula_dict_no_filter[context] = context_value['dim']

    return formula_dict, formula_dict_no_filter



formula_dict = create_formula_dict(data)[0]
formula_dict_no_filter = create_formula_dict(data)[1]

#replace alias in formula with source table/field
def replace_formula(formula_dict,formula):
    alias_list = re.findall(r'\[\w\w\w\w\w-\w\w\w\]', formula)
    for index,index_value in enumerate(alias_list):
        formula = formula.replace(alias_list[index],str(formula_dict[alias_list[index]]))
    return formula


rows_list.append(["Dashboard Title:",data['title']])
rows_list.append(["Data Source:",data['datasource']['title']])

#Make list of dashboard filters
if data['defaultFilters'] is not None:
    for filtr,filtr_value in enumerate(data['defaultFilters']):
        #to accomodate level style filters like Region>Zone>District>store
        if 'levels' in filtr_value:
            del filtr_value['model']['instanceid']
            dashboard_filter_name = filtr_value['model']
        else:
            filtr_dict = {key:val for key, val in filtr_value['jaql']['filter'].items() if (key != 'explicit' and key != 'multiSelection')}
            dashboard_filter_name = filtr_value['jaql']['dim'] + ' ' + str(filtr_dict)
        dashboard_filter_list.append(dashboard_filter_name)

rows_list.append(dashboard_filter_list)

rows_list.append(blank_cell)
rows_list.append(['Widget','Type','Panel Name','Source/Formula','Formula Details'])
   
    
#Loop through widgets
for widget,widget_value in enumerate(data['widgets']):
    widget_title = widget_value['title']
    if 'subtype' in widget_value:
        widget_type = widget_value['subtype']
    else:
        widget_type = widget_value['type']
        
    #check if widget already in list
    # if widget_title not in widget_title_list:
    widget_title_list.append(widget_title)
    rows_list.append([widget_title,widget_type])
        
    #within each widget, loop through panels
    for panel,panel_value in enumerate(widget_value['metadata']['panels']):

        panel_name = panel_value['name']

        #within each panel, loop through items and assign component to equal formula if it exists, else dim
        for item,item_value in enumerate(panel_value['items']):
            if panel_name == 'filters' and 'formula' in item_value['jaql']:
                filter_dict = {key:val for key, val in item_value['jaql']['filter'].items() if (key != 'explicit' and key != 'multiSelection' and key != 'by')}
                component = [item_value['jaql']['title'],filter_dict]
                formula_details = blank_cell
            elif 'formula' in item_value['jaql']:
                component = item_value['jaql']['title']
                formula_details = replace_formula(formula_dict,item_value['jaql']['formula'])
                #create formula summary list without filter
                newformula = replace_formula(formula_dict_no_filter,item_value['jaql']['formula'])
                newformula = newformula.replace(',','')
                if [component,newformula] not in formula_list:
                    formula_list.append([component,newformula])
            elif panel_name == 'filters':
                filter_dict = {key:val for key, val in item_value['jaql']['filter'].items() if (key != 'explicit' and key != 'multiSelection' and key != 'by')}
                component = [item_value['jaql']['dim'],filter_dict]
                formula_details = blank_cell
            elif 'agg' in item_value['jaql']:
                component = item_value['jaql']['title']
                formula_details = str(item_value['jaql']['agg']) + item_value['jaql']['dim']
                if [component,formula_details] not in formula_list:
                    formula_list.append([component,formula_details])
            else:
                component = item_value['jaql']['dim']
                formula_details = blank_cell
            
            rows_list.append([blank_cell,blank_cell,panel_name,component,formula_details])

filename_out = "%s.xlsx" % name

df1=pd.DataFrame(rows_list)
df2=pd.DataFrame(formula_list,columns=['Formula','Formula Details'])

# Create a new excel workbook
writer = pd.ExcelWriter(filename_out, engine='xlsxwriter')

# Write each dataframe to a different worksheet.
df1.to_excel(writer, sheet_name='Overview', index = False,header=None)
df2.to_excel(writer, sheet_name='Formula Summary', index = False)

writer.save()

# # Closing file
# f.close()
