# GoogleAPI Git

How to use Google Sheet and Google Drive API 

# Function

## Quickstart
    # Put gdapi.py into Lib folder
    import gdapi
    
    # Drive_1=gdapi.Drive(r'D:\Python\All_Practice\GoogleAPI\token\mnbbtoken.pickle')
    # Before ues API , log in your drive first.
    # Login Function:
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
## Create new Sheet
    Drive.create_newsheet('sheetname')    
## Create new doc
    Drive.create_doc('Docname')    
## Create new Slide
    Drive.create_Slides('Slidename')
## Create new Form
    Drive.create_form('Formname')
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
*```Before use the sheet api , you need to login first ( >>Drive=gdapi.Drive("PICKLE PATH OR JSON PATH") )```*
 
    sheet=gdapi.S(sheetid)  # if no sheetid it will create a newsheet 
   ## Sheet create
    sheet=gdapi.Sheet(name="NewTitle")  # it will create a new create which is named NewTitle 

   ## Write function  
    Data='Hi'
    sheet.write("A1",Data)  # A1='Hi'
    Data=[['Hi','I'],['am','Test']]
    sheet.write("A1:B2",Data)    # A1='Hi'  |  B1='I'
                                 # A2='am'  |  B2='Test'
    
   ## Set 1 behind Data argument 
    sheet.write("A1:B2",Data,1)  # A1='Hi'  |  B1='am'
                                 # A2='I'   |  B2='Test'
    Data=[1,2,3,4,5,6]                           
    sheet.write("A1:F1",Data)    # A1=1  |  B1=2 |  C1=3 |  D1=4 |  E1=5 |  F1=6 |
    
    
   ## Write across row data

    arr=['abc','def','g','ad','ee','bb']

    for i in range(0,len(arr)):
        sheet.write("A"+str(i),arr[i])
        # A1 = abc
        # A2 = def
        # A3 = g
        # A4 = ad

   ## Read sheet content
    
    string=sheet.read("A1")                      # return string
    list_value=sheet.read("workbook2!:A1:B3")    # return Value list

    
   ## Add new workbook in exist spreadsheet
    sheet.add_sheet("workbook1","workbook2")
    
   ## Get spreadsheet name
    sheet.title # return spreadsheet name
    
   ## Get spreadsheet id
    sheet.id # return sheet id
    
   ## Get spreadsheet url
    sheet.url # return sheet url
    
   ## Get spreadsheet sub
    sublist=sheet.sub # return sheet sub list
    
   ## Rename spreadsheet
    sheet.title="newtitle"
    
   ## Rename workbookdsheet
    sheet.resubtitle("oldtitle","newtitle")

