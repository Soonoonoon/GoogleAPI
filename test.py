import gdapi
import time
# Drive_1=gdapi.Drive(r'D:\Python\All_Practice\GoogleAPI\token\mnbbtoken.pickle')
Drive_1=gdapi.Drive("PICKLE PATH OR JSON PATH") # you can set another Drive2 ,etc...

# if you want to change Drive_1 variable to use another Credential
# use >> Drive_1.chose_json(path) or Drive_1.chose_pickle(path)
# if you want to create a new sheet

## Change Crendential
    # You can change different Drive to use or use more than 1 drive at the same time
    
    Drive.chose_json(jsonpath)
    Drive.chose_pickle(picklepath)    
## Download
    Drive.download("filename","destination of computer")
## Upload
    Drive.upload("filename","destination of drive")
## Delete
    Drive.delete("filename")
## Empty trash
    Drive.emptytrash()
## Find file
    match,similar=Drive_1.find_file(filename) # return 2 list, match = equal filename , similar= similar to filename
## Find folder
    folder_id=Drive_1.find_folder_id(folder)  # return Folder id
## Create new sheet
    Drive.create_newsheet('sheetname')
## Change permissions
    Link=Drive.change_permissions(file_id) # change permission to anyone can write and read , return a webViewLink
## Get webViewLink  
    webViewLink=Drive.get_weblink(file_id)


    
## Sheet
    sheet=Drive.Writer(sheetid)  # if no sheetid it will create a newsheet 
    
    Data='Hi'
    sheet.write("A1",Data)  # A1='Hi'
    Data=[['Hi','I'],['am','Test']]
    sheet.write("A1:B2",Data)    # A1='Hi'  |  B1='I'
                                 # A2='am'  |  B2='Test'
    
    #if set 1 behind Data 
    sheet.write("A1:B2",Data,1)  # A1='Hi'  |  B1='am'
                                 # A2='I'   |  B2='Test'
    Data=[1,2,3,4,5,6]                           
    sheet.write("A1:F1",Data)    # A1=1  |  B1=2 |  C1=3 |  D1=4 |  E1=5 |  F1=6 |

    # If write across row data

    arr=['abc','def','g','ad','ee','bb']

    for i in range(0,len(arr)):
        sheet.write("A"+str(i),arr[i])
        # A1 = abc
        # A2 = def
        # A3 = g
        # A4 = ad

    
    # Add new workbook in exist spreadsheet
    sheet.add_sheet("workbook1","workbook2")
    
    # get spreadsheet name
    sheet.title # return spreadsheet name
    
    # get spreadsheet id
    sheet.id # return sheet id
    
    # get spreadsheet url
    sheet.url # return sheet url
    
    # get spreadsheet sub
    sublist=sheet.sub # return sheet sub list
    
    # rename spreadsheet
    sheet.title="newtitle"
    
    # rename workbookdsheet
    sheet.resubtitle("oldtitle","newtitle")
