# GoogleAPI Git

How to use Google Sheet and Google Drive API 

# Function

## Quickstart
    # Put gdapi.py into Lib folder
    import gdapi
    Drive=gdapi.Drive(jsonpath or picklepath)
## Change Crendential
    # You can change different Drive to use or use more than 1 drive at the same time
    
    Drive.chose_json(jsonpath)
    Drive.chose_pickle(picklepath)
## Set download path
    Drive.Download_path=yourpath               # set download path, default is argvdirpath
## Download
    Drive.download("filename","destination of computer")
## Upload
    Drive.upload("filename","destination of drive")
## Delete
    Drive.delete("filename")
## Delete Emptyfolder
    Drive.delete_emptyfolder()
## Empty trashcan
    Drive.emptytrash()
## List all file
    Drive.list_all_file(view=1)# (self,*args,**kwargs) parameters: view=0/1 (list the found file , default 1)
## Find file
    dict_of_find=Drive_1.find_file(filename)   # return a dict
## Find folder
    folder_id=Drive_1.find_folder_id(folder)   # return Folder id
## Create new sheet
    Drive.create_newsheet('sheetname')
## Change permissions
    Link=Drive.change_permissions(file_id)     # change permission to anyone can write and read , return a webViewLink
## Get webViewLink  
    webViewLink=Drive.get_weblink(file_id)     # return weblink (string) 
## Get size
    size=Drive.get_size(file_id)               # return size    (string) 
## Get mimetype
    type_=Drive.get_type(file_id)              # return type of file
## Get LastModifyingUser
    user=Drive.get_lastModifyingUser(file_id)  # return dict of user
## Get shared state
    sharestate=Drive.get_shared(file_id)       # return boolean True / False
## Get Filelist
    Drive.get_filelist()                       # will write a list csv to default path
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
    
