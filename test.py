import gdapi
import time


# if you want to create a new sheet

file_id=gdapi.create_newsheet("new_sheet_name") # get a name of new_sheet_name Sheet

# change permission

gdapi.change_permissions(file_id) # 

# Write Sheet
sheet=gdapi.Writer(file_id)

sheet.write('B5','In new_sheet_name, Hi this is B5')
sheet.write("workbookname!A1:L1",['xuyy','xuyy','xuyy','xuyy','xuyy','xuyy',])#workbookname namechose

# Write lots of
arr=['abc','def','g','ad','ee','bb']
sheet.write("A1:L1",arr) 

# If write across row data

arr=['abc','def','g','ad','ee','bb']

for i in range(0,len(arr)):
    sheet.write("A"+str(i),arr[i])
    # A1 = abc
    # A2 = def
    # A3 = g
    # A4 = ad

    
# If write across row data method2

arr=[['abc','def','g'],['ad','ee','bb']]


sheet.write("A1:C2",arr)
    # A1 = abc  B1=def C1=g
    # A2 = def  B2=ee  C2=bb


# if write down in one column
sheet.write('A2:A3',['abb','cc'],1) # set 1 behind data = wirte down along column
        
# Create sheet method 2
sheet=gdapi.Writer
sheet.create("filename","workbook1","workbook2")# create a [filename ]sheet, workbook1 , workbook2, will under the page ,be newpage

# add new workbook in exist sheet
sheet.add_sheet("workbook1","workbook2")

# rename spreadsheet

sheet.title="newtitle"

# rename workbookdsheet

sheet.resubtitle("oldtitle","newtitle")

