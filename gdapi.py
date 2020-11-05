from __future__ import print_function
import pickle
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.http import MediaFileUpload, MediaIoBaseDownload
import io
import re
import datetime
# If modifying these scopes, delete the file token.pickle.

SCOPES = ['https://www.googleapis.com/auth/drive']
#mimetype dict
mimetype_dict={'xls': 'application/vnd.ms-excel', 'xlsx': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', 'xml': 'text/xml', 'ods': 'application/vnd.oasis.opendocument.spreadsheet', 'csv': 'text/csv', 'tmpl': 'text/plain', 'pdf': 'application/pdf', 'php': 'application/x-httpd-php', 'jpg': 'image/jpeg', 'png': 'image/png', 'gif': 'image/gif', 'bmp': 'image/bmp', 'txt': 'text/plain', 'doc': 'application/msword', 'js': 'text/js', 'swf': 'application/x-shockwave-flash', 'mp3': 'audio/mpeg', 'zip': 'application/zip', 'rar': 'application/rar', 'tar': 'application/tar', 'arj': 'application/arj', 'cab': 'application/cab', 'html': 'text/html', 'htm': 'text/html', 'default': 'application/octet-stream', 'folder': 'application/vnd.google-apps.folder', '': 'application/vnd.google-apps.video', 'Google Docs': 'application/vnd.google-apps.document', '3rd party shortcut': 'application/vnd.google-apps.drive-sdk', 'Google Drawing': 'application/vnd.google-apps.drawing', 'Google Drive file': 'application/vnd.google-apps.file', 'Google Drive folder': 'application/vnd.google-apps.folder', 'Google Forms': 'application/vnd.google-apps.form', 'Google Fusion Tables': 'application/vnd.google-apps.fusiontable', 'Google Slides': 'application/vnd.google-apps.presentation', 'Google Apps Scripts': 'application/vnd.google-apps.script', 'Shortcut': 'application/vnd.google-apps.shortcut', 'Google Sites': 'application/vnd.google-apps.site', 'Google Sheets': 'application/vnd.google-apps.spreadsheet'}
def format_str(namelimit,name):
	name_len=namelimit
	try:
		name_len=len(name.encode('big5'))
	except:
		pass
	fill_name=namelimit-name_len
	name_=name+(' '.encode('big5')*fill_name).decode('big5')
	return name_
def size_byte(size):
        dict_={1:1024,2:1024**2,3:1024**3,4:10**12}
        dict_unit={1:'KB',2:'MB',3:'GB',4:'TB'}
        for i in range(4,0,-1):
          if size>dict_[i]:
              return str(size//dict_[i])+'.'+str(size%dict_[i])[:2]+' '+dict_unit[i]
          if i==1:
              return str(size)+' Byte'
class Drive:
    
    PickleFile=''
    json_path=''
    def __init__(self,*args):
        if args:
                        if '.json' in str(args):
                                self.chose_json(args[0])
                        elif '.pickle' in str(args):
                                self.chose_pickle(args[0])
                        else:
                            service,service_sheet=self.main()
    def chose_json(self,path):
        global json_path,service,service_sheet
        json_path=path
        service,service_sheet=self.main()
        self.json_path=path
    def chose_pickle(self,path):
            global PickleFile,service,service_sheet
            PickleFile=path
            self.PickleFile=path
            service,service_sheet=self.main()
    def emptytrash(self):
        service.files().emptyTrash().execute()        
    def find_folder(self):
                results = service.files().list(q="mimeType='application/vnd.google-apps.folder'" and "trashed=false",
                                              spaces='drive',
                                              fields='nextPageToken, files(id, name)',
                                              
                                              pageToken=None).execute()

    def find_file(self,filename,*args):
          show_=1
          if args:
            show_=0  
          results = service.files().list(     q="trashed=false",
                                              pageSize=1000,
                                              fields='files(id, name)',
                                              pageToken=None).execute()
          
          
          items = results.get('files', [])
         
          
          if not items:return 0
          itemfind=[]
          similar=[]
          dict_of_all={}
          for item in items:
                 id_=item['id']
                 name_=item['name']
                 
                 if filename==name_:
                     itemfind.append(id_)
                 else:
                     if re.findall(filename,name_,re.IGNORECASE):
                         similar.append(id_)
          if itemfind :
            if show_:
              print("找到符合名稱的項目:")
            for i in itemfind:
              filedict=service.files().get(fileId=i).execute()
              name_=filedict['name']
              mimetype=filedict['mimeType']
              dict_of_all[i]=name_,mimetype
              if show_:print('ID: '+i+" | Name: {:<20}".format(name_))
          if similar :
           if show_:print("找到相似名稱的項目:")
           for i in similar:
              filedict=service.files().get(fileId=i).execute()
              name_=filedict['name']
              mimetype=filedict['mimeType']
              dict_of_all[i]=name_,mimetype
              if show_:print('ID: '+i+" | Name: {:<20}".format(name_))
          if not itemfind and not similar:
              print("找不到項目")
              return 0,0
          list_=sorted(dict_of_all.items(),key=lambda b:b[1],reverse=True)
          
          dict_of_all={}
          for i in list_:
              key,filename_filetype=i
              filename,filetype=filename_filetype
              dict_of_all[key]=filename,filetype
          
          
          
          return dict_of_all
    def listdir(self,fileid):
        response=service.files().list(q="'"+str(fileid)+"' in parents",
                                          fields='files(id, name)',
                                          ).execute()
        return response['files']
    def download_id(self,fileid,dst,*args):
        
        if fileid:
            
            filedict=service.files().get(fileId=fileid).execute()
            name_=filedict['name']
            mimeType_=filedict['mimeType']
            if not args:
                print("Download " +name_ )
            if 'folder'in str(mimeType_):
                list_of_files=self.listdir(fileid)
                fold_name=dst+'\\'+name_
                if not os.path.isdir(fold_name):
                    os.mkdir(fold_name)
                for each in list_of_files:
                    
                    self.download_id(each['id'],fold_name,1)
                print("Download 100%")
                return
            if 'application/vnd.google-apps' in str(mimeType_):
                if 'spreadsheet' in str(mimeType_):
                    mimeType_='text/csv'
                    name_=name_+'.csv'
                elif 'document'in str(mimeType_):
                    mimeType_='application/msword'
                    name_=name_+'.docx'
                
                request = service.files().export_media(fileId=fileid,mimeType=mimeType_)
            else:
                request = service.files().get_media(fileId=fileid)
            fh = io.BytesIO()
            downloader = MediaIoBaseDownload(fh, request)
            
            
            done = 0
            while not done :
                status, done = downloader.next_chunk()
                if not args:
                  print("Download %d%%" % int(status.progress() * 100))
            
            filepath=dst+'\\'+name_
            with io.open(filepath,'wb') as f:
                fh.seek(0)
                f.write(fh.read())
    def download(self,file,dst):
        if not dst:
            print("dstpath not found")
            return

        dict_of_find=self.find_file(file,1)
        temp_dict={}
        fileid=0
        count=0
        alreadyprint=[]
        if dict_of_find:
                        for i in dict_of_find:
                            if i in alreadyprint:continue
                            name_,mimetype_=dict_of_find[i]
                            count+=1
                            if 'folder' in str(mimetype_):
                                list_files=self.listdir(i)
                                size=self.get_size(i)
                                print(format_str(3,str(count)+'.'),name_+"  (Folder) :"+"( Total size: "+size+" )")
                               
                                temp_dict[str(count)]=i,name_
                                count+=1
                                for each in list_files:
                                    if each['id'] in dict_of_find:
                                        
                                        size=self.get_size(each['id'])
                                        print('╘═'+format_str(3,str(count)+'.'),format_str(40,each['name']),'| Size:',size)
                                        alreadyprint.append(each['id'])
                                        temp_dict[str(count)]=each['id'],each['name']
                                        count+=1
                            else:
                                size=self.get_size(i)
                                print(format_str(3,str(count)+'.'),format_str(40,name_),'| Size:',size)
                            
                                temp_dict[str(count)]=i,name_
        print("\nIf wnat to download:  1. %s  >> press 1:"%(temp_dict[str(1)][1]),'(* press enter to skip ) ')
        chose=input('[* Chose more than one : press 1~3(chose: 1,2,3) or 1,3,5(chose: 1,3,5)]\n')
        
        if str(chose) in temp_dict:
                fileid,name_s=temp_dict[str(chose)]
        elif chose:
            if '~' in str(chose):
                file_=chose.split('~')
             
                for chose_ in range(int(file_[0]),int(file_[-1])+1):
                    if str(chose_) in temp_dict:
                        fileid,name_s=temp_dict[str(chose_)]
                
                        self.download_id(fileid,dst)
                     
                return    
            elif ',' in str(chose):
                file_=chose.split(',')
                for chose_ in file_:
                    if str(chose_) in temp_dict:
                        fileid,name_s=temp_dict[str(chose_)]
                        self.download_id(fileid,dst)
                return
                
        if fileid and 'list' not in str(type(fileid)):
            
            filedict=service.files().get(fileId=fileid).execute()
            name_=filedict['name']
            mimeType_=filedict['mimeType']
            
            print("Download " +name_ )
            if 'folder'in str(mimeType_):
                list_of_files=self.listdir(fileid)
                fold_name=dst+'\\'+name_
                if not os.path.isdir(fold_name):
                    os.mkdir(fold_name)
                for each in list_of_files:
                    
                    self.download_id(each['id'],fold_name,1)
                print("Download 100%")
                return
            if 'application/vnd.google-apps' in str(mimeType_):
                if 'spreadsheet' in str(mimeType_):
                    mimeType_='text/csv'
                    name_=name_+'.csv'
                elif 'document'in str(mimeType_):
                    mimeType_='application/msword'
                    name_=name_+'.csv'
                
                request = service.files().export_media(fileId=fileid,mimeType=mimeType_)
            else:
                request = service.files().get_media(fileId=fileid)
            fh = io.BytesIO()
            downloader = MediaIoBaseDownload(fh, request)
            
            
            done = 0
            while not done :
                status, done = downloader.next_chunk()
                print("Download %d%%" % int(status.progress() * 100))
            
            filepath=dst+'\\'+name_
            with io.open(filepath,'wb') as f:
                fh.seek(0)
                f.write(fh.read())
    def change_permissions(self,file_id):# 將某ID的檔案權限更改
        service.permissions().create(fileId=file_id,
                                             body= {
                                            'role': 'writer',
                                            'type': 'anyone',
                                              }
                                            ).execute()
    def create_newsheet(self,*filename):
        if not filename:
            filename="NewSheet"
        file_metadata = {
        'name': filename,
        'mimeType': 'application/vnd.google-apps.spreadsheet'
        }
        
        file = service.files().create(body=file_metadata).execute()
        
        return file['id']
    def create_doc(self,*filename):
        if not filename:
            filename="NewDoc"
        file_metadata = {
        'name': filename,
        'mimeType': 'application/vnd.google-apps.document'
        }
        
        file = service.files().create(body=file_metadata).execute()
        
        return file['id']
    def create_Slides(self,*filename):
        if not filename:
            filename="NewSlide"
        file_metadata = {
        'name': filename,
        'mimeType': 'application/vnd.google-apps.presentation'
        }
        
        file = service.files().create(body=file_metadata).execute()
        
        return file['id']
    def create_form(self,*filename):
        if not filename:
            filename="NewSlide"
        file_metadata = {
        'name': filename,
        'mimeType': 'application/vnd.google-apps.form'
        }
        
        file = service.files().create(body=file_metadata).execute()
        
        return file['id']
    def find_folder_id(self,foldername):
          results = service.files().list(q="mimeType='application/vnd.google-apps.folder'"and "trashed=false",
                                      
                                              spaces='drive',
                                              
                                              fields='nextPageToken, files(id, name)',
                                              pageToken=None).execute()
          items = results.get('files', [])
          
          if not items:return 0
          
          for item in items:
                 id_=item['id']
                 name_=item['name']
                 
                 if foldername==name_:
                     #folderID=id_
                     return id_
          return 0
    def upload(self,filepath,*dstpath):# 上傳檔案到特定資料夾
        filename=''
        foldername=''
        if os.path.isdir(filepath) :
            filepath=os.path.abspath(filepath)
            foldername=os.path.basename(filepath)
            folderID=''
            if dstpath:
                
                foldername=os.path.dirname(dstpath[0])
                if not foldername:
                    foldername=dstpath[0]
                folderID=self.find_folder_id(foldername)
                if folderID:
                    file_metadata = {
                    'name': foldername,
                    'parents' : [folderID],
                    'mimeType': 'application/vnd.google-apps.folder'
                    }
                else:
                     
                     file_metadata = {
                    'name': foldername,
                   
                    'mimeType': 'application/vnd.google-apps.folder'
                     }

                     file = service.files().create(body=file_metadata).execute()
                     folderID=file['id']
                     
            
            if not folderID:
                
                
                folderID=self.find_folder_id(foldername)
              
                
                if not folderID:
                    file_metadata = {
                        'name': foldername,
                        
                        'mimeType': 'application/vnd.google-apps.folder'
                        }
                    file = service.files().create(body=file_metadata).execute()
                    
                else:
                    file_metadata = {
                    'name': foldername,
                    'parents' : [folderID],
                    'mimeType': 'application/vnd.google-apps.folder'
                    }
            
            for i in os.listdir(filepath):
                eachnewpath=filepath+'/'+i
                
                if foldername:# 代表有指定資料夾
                    dstpath=foldername+'/'+i
                    
                    print("Uploading: ",i)
                    
                    
                   
                    self.upload(eachnewpath,dstpath)
            return
        if dstpath:
        
            foldername=os.path.dirname(dstpath[0])
            filename=os.path.basename(dstpath[0])
        if not filename:
               filename=os.path.basename(filepath)
        if '.' in str(os.path.splitext(filename)[-1]):
            Extension=os.path.splitext(filename)[-1].replace('.','')
            if Extension in mimetype_dict:
                file_type=mimetype_dict[Extension]
            else:
                file_type='application/octet-stream'
        else:
            print("Mimetype not found , try to upload this folder")
            foldername=os.path.basename(dstpath[0])
            filename=os.path.basename(filepath)
            Extension=os.path.splitext(filename)[-1].replace('.','')
            
            if Extension in mimetype_dict:
                
                file_type=mimetype_dict[Extension]
            else:file_type='application/octet-stream'
            

        
        
        
       
        if foldername:
          
          folderID=0
          folderID=self.find_folder_id(foldername)
          
         
          if  not folderID:
              file_metadata = { 'name': filename}
          else:
              file_metadata = {
                'name': filename,
                'parents': [folderID]
                }       
        else:
              file_metadata = { 'name': filename}          
        
        
        media = MediaFileUpload(filepath,
                                mimetype=file_type,
                                resumable=True)
        file = service.files().create(body=file_metadata,
                                            media_body=media,
                                            fields='id').execute()
        print("Upload Finish !")
    def delete_id(self,fileid):
        file_=service.files().get(fileId=fileid).execute()
        print("Delete: ",file_['name'])
        try:
            service.files().delete(fileId=fileid).execute()
        except :
          print("Delete Error")
    def delete(self,filename):
      dict_of_find=self.find_file(filename,1)
      temp_dict={}
      fileid=0
      count=0
      alreadyprint=[]
      if dict_of_find:
                    for i in dict_of_find:
                        if i in alreadyprint:continue
                        name_,mimetype_=dict_of_find[i]
                        count+=1
                        if 'folder' in str(mimetype_):
                            size=self.get_size(i)
                            list_files=self.listdir(i)
                            print(format_str(3,str(count)+'.'),name_+"  (Folder) :"+" (Total size: "+size+")")
                            
                            temp_dict[str(count)]=i,name_
                            count+=1
                            for each in list_files:
                                if each['id'] in dict_of_find:
                                    size=self.get_size(each['id'])
                                    print('╘═'+format_str(3,str(count)+'.'),format_str(40,each['name']),'| Size:',size)
                                    alreadyprint.append(each['id'])
                                    temp_dict[str(count)]=each['id'],each['name']
                                    
                                    count+=1
                            
                        else:
                            size=self.get_size(i)
                            print(format_str(3,str(count)+'.'),format_str(40,name_),'| Size:',size)
                        
                            temp_dict[str(count)]=i,name_
      
      print("\nIf wnat to delete:  1. %s  >> press 1:"%(temp_dict[str(1)][1]),' (* press enter to skip )')
      chose=input('[* Chose more than one : press 1~3(chose: 1,2,3) or 1,3,5(chose: 1,3,5)]\n')
      if str(chose) in temp_dict:
                fileid,name_s=temp_dict[str(chose)]
      elif chose:
        if '~' in str(chose):
            file_=chose.split('~')
            for chose_ in range(int(file_[0]),int(file_[-1])+1):
                if str(chose_) in temp_dict:
                    fileid,name_s=temp_dict[str(chose_)]
                    
                    self.delete_id(fileid)
                 
            return    
        elif ',' in str(chose):
            file_=chose.split(',')
            for chose_ in file_:
                if str(chose_) in temp_dict:
                    fileid,name_s=temp_dict[str(chose_)]
                    self.delete_id(fileid)
            return  
      if not fileid:
          
          return
      
      """Permanently delete a file, skipping the trash

      Args:
        service: Drive API service instance.
        file_id: ID of the file to delete.
      """
      file_=service.files().get(fileId=fileid).execute()
      print("Delete: ",file_['name'])
      try:
            service.files().delete(fileId=fileid).execute()
      except :
          print("Delete Error")
    def get_weblink(self,fileid):
        # fields = get what you want
        # file resource url=https://developers.google.com/drive/api/v3/reference/files
        results = service.files().get(fileId=fileid,fields='webViewLink').execute()['webViewLink']
        return results
    
    def get_size(self,fileid,*arg):
        dict_={1:1024,2:1024**2,3:1024**3,4:10**12}
        dict_unit={1:'KB',2:'MB',3:'GB',4:'TB'}
        type_=self.get_type(fileid)
        if 'folder' in str(type_):
            sizes=0
            list_files=self.listdir(fileid)
            for j in list_files:
                size=self.get_size(j['id'],0)
                sizes+=int(size)
            return size_byte(sizes)
        results = service.files().get(fileId=fileid,fields='size').execute()['size']
        if arg:  return results
        return size_byte(int(results))
    def get_type(self,fileid):
        # fields = get what you want
        # file resource url=https://developers.google.com/drive/api/v3/reference/files
        results = service.files().get(fileId=fileid,fields='mimeType').execute()['mimeType']
        return results
    def get_lastModifyingUser(self,fileid):
        # fields = get what you want
        # file resource url=https://developers.google.com/drive/api/v3/reference/files
        results = service.files().get(fileId=fileid,fields='lastModifyingUser').execute()['lastModifyingUser']
        return results
    def get_shared(self,fileid):
        # fields = get what you want
        # file resource url=https://developers.google.com/drive/api/v3/reference/files
        results = service.files().get(fileId=fileid,fields='shared').execute()['shared']
        return results
    def main(self):
        """Shows basic usage of the Drive v3 API.
        Prints the names and ids of the first 10 files the user has access to.
        """
        creds = None
        # The file token.pickle stores the user's access and refresh tokens, and is
        # created automatically when the authorization flow completes for the first
        # time.
        if os.path.exists(PickleFile):
            with open(PickleFile, 'rb') as token:
                creds = pickle.load(token)
        # If there are no (valid) credentials available, let the user log in.
        try:
         if not creds or not creds.valid:
            if creds and creds.expired and creds.refresh_token:
                creds.refresh(Request())
            else:
                flow = InstalledAppFlow.from_client_secrets_file(
                    json_path, SCOPES)
                creds = flow.run_local_server(port=8012)
            # Save the credentials for the next run
            with open(PickleFile, 'wb') as token:
                pickle.dump(creds, token)
        except:
            print("▲ If you don't set the json or pickle path, plz set the path before use ")
            print("Use function: \n 1. chose_pickle (pickle_path)\n 2. chose_json   (json_path)")
            return
        service = build('drive', 'v3', credentials=creds)
        service_sheet = build('sheets', 'v4', credentials=creds)
             
     
        # Call the Drive v3 API#
       
        
        return service,service_sheet
    def sheet_update(self,sheetid,workRange,content):
        if 'list' not in str(type(content)):
                content=[content]
        writeData = {}
       
       
        writeData['values']= content
        writeData['majorDimension']="ROWS"
        writeData['range']= workRange
        
        
        request = service_sheet.spreadsheets().values().update(spreadsheetId=sheetid, range=workRange,valueInputOption='RAW',body=writeData )
        response = request.execute()   
    def sheet_append(self,sheetid,workRange,content):
        if 'list' not in str(type(content)):
                content=[content]
        writeData = {}
        
     
        writeData['values']= content
        writeData['majorDimension']="ROWS"
        writeData['range']= workRange
        
       
        
        request = service_sheet.spreadsheets().values().append(spreadsheetId=sheetid, range=workRange,valueInputOption='RAW', insertDataOption='INSERT_ROWS', body=writeData )
        response = request.execute()
    
    try:
        service,service_sheet = main()
    except:
        pass

class Writer:
        
        def __init__(self,*id_in):
        
            global id_
            if id_in:
               self._id=id_in[0]
               id_=id_in[0]
               self.getsheet()
            else:
                timenow=datetime.datetime.now().strftime("%H%M%S")
                self._id=self.create("NewSheet"+str(timenow))
                id_=self._id
                print("Can't find SheetID , automake  NewSheet"+str(timenow))
           
        @property
        def id(self):
            return self._id
        @property
        def url(self):
            return self._url
        @property
        def properties(self):
            return self._properties
        @property
        def sub(self):
            return self._sub
        def getsheet(self):
                global id_
                
                request = service_sheet.spreadsheets().get(spreadsheetId=id_, includeGridData=False)
                response=request.execute()
                self._sub={}
                for i in response['sheets']:
                    eachtitle=i['properties']['title']
                    eachid=i['properties']['sheetId']
                    eachindex=i['properties']['index']
                    self._sub[eachindex]=eachid,eachtitle
                    #self._sub[eachindex].title= eachtitle
                    #self._sub[eachindex].id   = eachid
                    #names = locals()
                    #self.+names[self+]=564
                self._title=response['properties']['title']
                self._url=response['spreadsheetUrl']
                self._properties=response['properties']
           #except Exception as Err:
             #   if 'was not found' in str(Err):return
        def add_sheet(newpagename,*args):   #create　new page in exist spreadsheet
            global id_
            request_=[{"addSheet":{"properties":{"title": newpagename}}}]
            if  args:
                for new_page in args:
                    request_.append({"addSheet":{"properties":{"title": new_page}}})
            body={"requests":request_}
            request = service_sheet.spreadsheets().batchUpdate(spreadsheetId=id_,body=body)
            request.execute()
        @property
        def title(self):
            return self._title
        @title.setter
        def title(self,newname):
            
            self._title=newname
            self.rename_sheet(newname)
        @property
        def subtitle(self):
            return self._title
        @title.setter
        def title(self,newname,*args):
            
            self._title=newname
            self.rename_sheet(newname)
        def resubtitle(self,newname,args): 
            if args:
                oldname=newname
                newname=args
                for j in self.sub:
                    sheetId,searname=self.sub[j]
                    if oldname==searname:
                       
                      self.rename_workbook(sheetId,newname)
                      return
        def rename_sheet(self,newtitle):
            body={"requests":[{
                "updateSpreadsheetProperties":{
                    "properties":{
                        "title": newtitle},
                        "fields": '*'}
                  }]
                 }
            request = service_sheet.spreadsheets().batchUpdate(spreadsheetId=id_ ,body=body)
            request.execute()
        def rename_workbook(self,sheetId,newname):
            body={"requests":[{
                "updateSheetProperties":{
                    "properties":{
                        "sheetId":sheetId,"title": newname},"fields": 'title'}}]}
            request = service_sheet.spreadsheets().batchUpdate(spreadsheetId=id_ ,body=body)
            request.execute()
        @classmethod    
        def create(cls,filename,*args):  #create　new spreadsheet
            global id_
            sheet_page=[]
            if  args:
                for new_page in args:
                    sheet_page.append({"properties":{"title": new_page}})
            body= {"properties": {"title":filename},"sheets":sheet_page }
            request = service_sheet.spreadsheets().create(body=body)
            response=request.execute()
            
            
            id_=response['spreadsheetId']
            cls.id=id_
            cls.getsheet(cls)
        
        def copy(self,dst):             # copy a new workbook in exist spreadsheet
            body={ 'destination_spreadsheet_id': self.id}
            
            request = service_sheet.spreadsheets().sheets().copyTo(spreadsheetId=self.id,sheetId=0,body=body)
            request.execute()
        def write(self,workRange,content,*args):
            
                
            writeData = {}
            if 'list' not in str(type(content)):
                content=[content]
                writeData['values']=[content]
            
                      
            
            
            elif 'list'in str(type(content)):
                  if isinstance(content[0],list):# represent this is 2 dimensional list
                       writeData['values']=content
                  else:
                       writeData['values']=[content]
            if args:
                writeData['majorDimension']="COLUMNS"
            else:
                writeData['majorDimension']="ROWS"
            writeData['range']= workRange
        
       
        
            request = service_sheet.spreadsheets().values().update(spreadsheetId=self.id, range=workRange,valueInputOption='RAW', body=writeData )
            response = request.execute()


 
