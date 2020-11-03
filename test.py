import gdapi

# if you want to create a new sheet

file_id=gdapi.create_newsheet("new_sheet_name") # get a name of new_sheet_name Sheet

# change permission

gdapi.change_permissions(file_id) # 

# Write Sheet
sheetid=file_id
gdapi.write(sheetid,'B5','In new_sheet_name, Hi this is B5')
