import gdapi
import time
# = = = =  Drive = = = =


# Drive=gdapi.Drive(r'D:\Python\GoogleAPI\token\your.pickle')
# Before ues API , log in your drive first.
# Login Function:
Drive=gdapi.Drive("PICKLE PATH OR JSON PATH") # you can set another Drive2 ,etc...
# json file accept : 
#       1. credentials json of OAuth2.0
#       2. keys json of service account
# if you want to change Drive variable to use another Credential
# use >> Drive.chose_json(path) or Drive.chose_pickle(path)

# Troubleshoot :
# If get "ERR_CONNECTION_REFUSED" when access 
# try : Use Private Browsing to login
if Drive:

    
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
## Move file
    Drive.move(Fileid,Foldername)
## Empty trashcan
    Drive.emptytrash()
## List all file
    Drive.list_all_file(view=1)# (self,*args,**kwargs) parameters: view=0/1 (list the found file , default 1)
## Find file
    dict_of_find=Drive.find_file(filename)   # return a dict
## Find folder
    folder_id=Drive.find_folder_id(folder)   # return Folder id
## Create NewFolder
    Drive.mkdir(foldername) 
## Create new Sheet
    Drive.create_newsheet('sheetname')

    #example :
    # 1.  Drive.create_newsheet('sheetname',folder='Fold') # It will find the folder , and create under the folder , or create a new folder
    # 2.  Drive.create_newsheet('sheetname','Fold')
    # example 1 equal 2
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


    
# = = = =  Sheet = = = =
    #Before use the sheet api , you need to login first ( >>Drive=gdapi.Drive("PICKLE PATH OR JSON PATH") )

    sheet=gdapi.Sheet(sheetid)  # if no sheetid it will create a newsheet

    # Sheet create
    sheet=gdapi.Sheet(name="NewTitle")  # it will create a new create which is named NewTitle
    # Sheet create example:
    sheet=gdapi.Sheet(name="NewTitle",folder="Foldername") # it will create a new create which is named NewTitle.
                                                           # and if find the folder name it will put sheet in
                                                           # or create a new folder


    # Write
    Data='Hi'
    sheet.write("A1",Data)       # A1='Hi'
    Data=[['Hi','I'],['am','Test']]
    sheet.write("A1:B2",Data)    # A1='Hi'  |  B1='I'
                                 # A2='am'  |  B2='Test'
    # Write with Cell :
    sheet.write("1:4",Data)
    #sheet.write("A1:B2",Data) == sheet.write("1:4",Data)



    # Write cell
    sheet.write("1:2",Data)
    #if set 1 behind Data 
    sheet.write("A1:B2",Data,1)  # A1='Hi'  |  B1='am'
                                 # A2='I'   |  B2='Test'
    Data=[1,2,3,4,5,6]                           
    sheet.write("A1:F1",Data)    # A1=1  |  B1=2 |  C1=3 |  D1=4 |  E1=5 |  F1=6 |
    
    
    # Write across row data

    arr=['abc','def','g','ad','ee','bb']

    for i in range(0,len(arr)):
        sheet.write("A"+str(i),arr[i])
        # A1 = abc
        # A2 = def
        # A3 = g
        # A4 = ad

    # Read sheet content
    
    string=sheet.read("A1")                      # return string
    list_value=sheet.read("workbook2!:A1:B3")    # return Value list

    # Delete content
    sheet.delete("A1:F1")    #A1~F1 will be deleted

    # Color cell
    sheet.setcolor("A1",(255,120,0),) # (RangeSetColor,Tuple(R,G,B),HexColor,Alpha=0.5,SheetId=0)
    #              Description                  |   Accept  Value
    #----------------------------------------------------------------
    #  RangeSetColor = set color on which cell  | Set  > "1:1"= "A1"
    #  Tuple(R,G,B)  = set color RGB            | Set  > (255,255,0) # accept tuple and int
    #　HexColor      = set color in Hex         | Set  > #CAFFFF    , if set Hexcolor , (R,G,B) can put anything except tuple.
    #  Alpha         = set transparent          | Set  > [0,1]      , default=1
    #  SheetId       = set sheetid              | Set  > long int   , default=0  , the number  in url gid=[sheetid] 

    # Color method example2
    sheet.setcolor("A1",0,'#CAFFFF',)

    # Color method example3 : change to workbook3 
    sheet.setcolor("A1",0,'#CAFFFF',sheetid=3)
                   
    # Add new workbook in exist spreadsheet
    sheet.add_sheet("workbook1","workbook2")
    
    # Get spreadsheet name
    sheet.title # return spreadsheet name
    
    # Get spreadsheet id
    sheet.id # return sheet id
    
    # Get spreadsheet url
    sheet.url # return sheet url
    
    # Get spreadsheet sub
    sublist=sheet.sub # return sheet sub list
    
    # Rename spreadsheet
    sheet.title="newtitle"
    
    # Rename workbookdsheet
    sheet.resubtitle("oldtitle","newtitle")

    # Adjust_column and row       
    Colrange="1:5" # Column A-D =1~4    * also accept: "A:D"
    Col_pixel= 20  # 20 pixels between  column 
    Rowrange="1:4" # Row= 1~3
    Row_pixel= 40  # 40 pixels between  row 
    sheetid=0      # default=0
    sheet.adjust_col_row(Colrange,Col_pixel,Rowrange,Row_pixel,*sheetid)
    # Only adjust column
    Colrange="1:5" # column 1-4    * also accept: "A:D"
    Col_pixel=40
    sheet.adjust_col(Colrange,Col_pixel,*sheetid)
    # Only adjust row
    Rowrange="2:4" # row 2-3
    Row_pixel=40
    sheet.adjust_row(Rowrange,Row_pixel,*sheetid)
        
    # Append Column
    sheet.append_col(length,*sheetid,**kwargs) #append length column
    # parameter:
    #   sheetid = 123456  # add a sheet id to access
    # kwargs:
    #   sheetname= 'abc'  # choose sheetname to access


    # Delete Column
    start=1
    end=6  
    sheet.append_col((start,end),*sheetid,**kwargs) # (start,end) is a tuple data
    # also can use alphabet to represent Column number
    sheet.append_col('a:e',*sheetid,**kwargs)       # delete column from a to e , a will keep 

    
    # Append Row
    sheet.append_row(length,*sheetid,**kwargs) #append length row
    


    # Delete Row
    start=1# 
    end=6  # 
    sheet.append_row((start,end),*sheetid,**kwargs) # (start,end) is a tuple data

    
    # Find and Replace
    find='string'           #   find     string
    reapcle='test'          #   replace string
    allsheet=0              #   0=find/replace only one page,1=find/replace all page | [Optional]
    
    sheet.FNR(find,replace,*sheetID,**kwargs)

    # example :
    id=123456789            #   must be positive number
    sheet.FNR(find,replace,id=id)
    # id set method2
    sheet.FNR(find,replace,id)
    
    # FNR by specify workbook name
    name="workbookname"     #   must be string
    sheet.FNR(find,replace,name=name)

    # FNR of all spreadsheet    
    sheet.FNR(find,replace,allsheet=1)
    
    # Sort Data
    colname="ABC"    #  sort by colname
    col=1            #  sort by column 1 
    UpOrDown=0       #  0/1  ( DESCENDING / ASCENDING )  ,Defalut 1
    #sheet.sort_col(colname,*UpOrDown,**col)
    sheet.sort_col('',0,col=col)    # DESCENDING and sort by col 1
    sheet.sort_col(colname,1)       # ASCENDING  and sort by colname

    # Use Formula
    
    workrange= "A1"   # put formula in which cell
    
    formula="sum(B1:B20)"  # Excel/Sheet Formula format 
    sheet.formula(workrange,formula,*sheetid)
     # If what to use match function: "match("+findwhat+","+findrange+",0)
    workrange= "A1" 
    findwhat="name"
    findrange="B1:B20"
    formula='match("'+findwhat+","+findrange+'",0)'
    sheet.formula(workrange,formula,*sheetid)

     # Use formula get result in different workbook
    sheet.formula("workbook2!B25",'sum(workbook1!B2:B5)')  # Get result in workbook2>B25 , execute function in wrokbook1 >B2:B25

    # Get col index
    findname="ABC"
    index_col=sheet.get_col_index(sheetname,*sheetID,**kwargs)  # return index of column
    # kwargs :
    # sheetname | example : sheetname='A' , find the column name in Sheetname  A

    # Get row index
    findname="ABC"
    index_row=sheet.get_col(sheetname,*sheetID)  # return index_row number

    # Get worksheet id
    sheet_id=sheet.getsub_id(sheetname)

    # Get worksheet size (max column and max row)
    sheet_max_column,sheet_max_row=sheet.get_sheet_size(sheetname)
    
    # Find string
    find_sting="string to find"
    find_string(find_sting,*sheetID,**kwargs)  # return a list of cell
        
