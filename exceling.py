# Import pandas
import pandas as pd
# from win32com.client import Dispatch

# Declarations
myFile = 'workbook.xlsx'
outputFile = 'output.xlsx'
NO_RULE = 'NO_RULE'

# Close Files
# xl = Dispatch('Excel.Application')
# xl.Workbooks.Open(myFile).close()
# xl.Workbooks.Open(outputFile).close()

# Load workbook
wb = pd.ExcelFile(myFile)
print("File: " + myFile + " loaded")

# Load in sheets to be data frames
DCSList = wb.parse('DCSList')
IOList = wb.parse('IOList')
equipList = wb.parse('EquipList')
rules = wb.parse('Rules')
PIConfigSettings = wb.parse('PIConfigSettings').T
print("Lists created")
print(PIConfigSettings)

# Join lists (SQL left joins)
myList = DCSList.merge(IOList, on='TagName', how='left')
myList = myList.merge(equipList, left_on='Parent', right_on='EquipNumber', how='left')
print("Lists joined")

# Add placeholder for applicable rule
myList['PIConfigSetting'] = NO_RULE

# Iterate through the main list
for myListIndex, myListRow in myList.iterrows():
    print "Processing Tag: " + myListRow['TagName']
    for ruleIndex, ruleRow in rules.iterrows():
        # Check if rule applies to TagName
        if ruleRow['Rule'] in myListRow['TagName']:
            print ruleRow['PIConfigSetting']
            # Add Rule to list and add settings
            myListRow['PIConfigSetting'] = ruleRow['PIConfigSetting']
            break

# Get PI Point List
includedTags = myList.loc[myList['PIConfigSetting'] != NO_RULE]

# Get tags not included
excludedTags = myList.loc[myList['PIConfigSetting'] == NO_RULE]

# Add config settings to included tags
includedTags = includedTags.merge(PIConfigSettings, on='PIConfigSetting', how='left')
print(PIConfigSettings)

# Specify a writer
writer = pd.ExcelWriter(outputFile, engine='xlsxwriter')

# Write your DataFrame to a file
includedTags.to_excel(writer, 'PITags')
print('Total Included Tags: ' + str(includedTags.shape[0]))
excludedTags.to_excel(writer, 'ExcludedTags')
print('Total Excluded Tags: ' + str(excludedTags.shape[0]))
print("Lists written to file: " + outputFile)

# Save the result
writer.save()
