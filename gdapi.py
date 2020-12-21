from __future__ import print_function
import pickle
import os.path
import requests
from oauth2client.service_account import ServiceAccountCredentials
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.http import MediaFileUpload, MediaIoBaseDownload
import io
import re
import datetime
from  pprint  import pprint
import sys
import time
import base64
import threading
from bs4 import BeautifulSoup
non_bmp_map = dict.fromkeys(range(0x10000, sys.maxunicode + 1), 0xfffd)
# If modifying these scopes, delete the file token.pickle.

SCOPES = ['https://www.googleapis.com/auth/drive']
Mail_SCOPES = ['https://mail.google.com/']
# Download path default
Download_path=os.path.split(sys.argv[0])[0]
#mimetype dict
mimetype_dict={'xls': 'application/vnd.ms-excel', 'xlsx': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', 'xml': 'text/xml', 'ods': 'application/vnd.oasis.opendocument.spreadsheet', 'csv': 'text/csv', 'tmpl': 'text/plain', 'pdf': 'application/pdf', 'php': 'application/x-httpd-php', 'jpg': 'image/jpeg', 'png': 'image/png', 'gif': 'image/gif', 'bmp': 'image/bmp', 'txt': 'text/plain', 'doc': 'application/msword', 'js': 'text/js', 'swf': 'application/x-shockwave-flash', 'mp3': 'audio/mpeg', 'zip': 'application/zip', 'rar': 'application/rar', 'tar': 'application/tar', 'arj': 'application/arj', 'cab': 'application/cab', 'html': 'text/html', 'htm': 'text/html', 'default': 'application/octet-stream', 'folder': 'application/vnd.google-apps.folder', '': 'application/vnd.google-apps.video', 'Google Docs': 'application/vnd.google-apps.document', '3rd party shortcut': 'application/vnd.google-apps.drive-sdk', 'Google Drawing': 'application/vnd.google-apps.drawing', 'Google Drive file': 'application/vnd.google-apps.file', 'Google Drive folder': 'application/vnd.google-apps.folder', 'Google Forms': 'application/vnd.google-apps.form', 'Google Fusion Tables': 'application/vnd.google-apps.fusiontable', 'Google Slides': 'application/vnd.google-apps.presentation', 'Google Apps Scripts': 'application/vnd.google-apps.script', 'Shortcut': 'application/vnd.google-apps.shortcut', 'Google Sites': 'application/vnd.google-apps.site', 'Google Sheets': 'application/vnd.google-apps.spreadsheet'}
# number   <==> alphabet 
num_alphabet = {0:'',1: 'a', 2: 'b', 3: 'c', 4: 'd', 5: 'e', 6: 'f', 7: 'g', 8: 'h', 9: 'i', 10: 'j', 11: 'k', 12: 'l', 13: 'm', 14: 'n', 15: 'o', 16: 'p', 17: 'q', 18: 'r', 19: 's', 20: 't', 21: 'u', 22: 'v', 23: 'w', 24: 'x', 25: 'y', 26: 'z'}
# alphabet <==> number
alphabet_num = {'a': 1, 'b': 2, 'c': 3, 'd': 4, 'e': 5, 'f': 6, 'g': 7, 'h': 8, 'i': 9, 'j': 10, 'k': 11, 'l': 12, 'm': 13, 'n': 14, 'o': 15, 'p': 16, 'q': 17, 'r': 18, 's': 19, 't': 20, 'u': 21, 'v': 22, 'w': 23, 'x': 24, 'y': 25, 'z': 26}
Month_num={'jan':1,'feb':2,'mar':3,'apr':4,'may':5,'jun':6,'jul':7,'aug':8,'sep':9,'oct':10,'nov':11,'dec':12}

def download_recursive(gd,size,dist,fileid,PickleFile,*arg,**kwargs):
        if kwargs:
                newgd=Drive(PickleFile)
                if 'index' in kwargs:
                   index=kwargs['index']
                   
                if 'last' in kwargs:
                    last=1
                    start=(index-1)*dist
                    end=size
                else:
                    start=(index-1)*dist
                    
                    end=dist*(index-1)+dist
                print(index, ' download')
                print(start,end)
                print('=====')
                request = newgd.service.files().get_media(fileId=fileid)
                request.headers["Range"] = "bytes={}-{}".format(start,end)
                fh = io.BytesIO()
                downloader = MediaIoBaseDownload(fh, request)

            
                done = 0
                while not done :
                    status, done = downloader.next_chunk()
                    print("Download %d%%" % int(status.progress() * 100),str(index))
                #fh = io.BytesIO(request.execute())
                
                gd.downloader[index-1]=fh.getvalue()
                print("OK",index)
                newgd.service.close()
                newgd.service_sheet.close()

        
def split_avaerage(num,dist):
        num=int(num)
        if num<dist:return 1
        times=num//dist
        
        if times<1:times=1
        return times
def download_thread(fileid,gd,dst,PickleFile):
            filedict=gd.service.files().get(fileId=fileid,fields='*').execute()
            name_=filedict['name']
            mimeType_=filedict['mimeType']
            ownedByMe_=filedict['ownedByMe']
            if 'false' in str(ownedByMe_):return
            if 'trashed' in filedict:
                trash_=filedict['trashed']
                if 'true' in str(trash_).lower().strip():return
            
            if 'folder'in str(mimeType_):
                fold_name=dst+'\\'+str(name_)
                if not os.path.isdir(fold_name):
                    os.mkdir(fold_name)
                    return
            
            size=filedict['size']
            set_split_size=1024*1024*1024*2 #1GB
            fh=''
            
            if int(size )>int(set_split_size):
                
                split_num_=split_avaerage(size,set_split_size)#分割成多少個
                print(split_num_,123456)
                gd.downloader=[0]*split_num_
                for i in range(1,split_num_+1):
                    time.sleep(2)
                    if i==split_num_:
                        
                        threading.Thread(target=download_recursive,args=(gd,size,set_split_size,fileid,PickleFile),kwargs={'last':1,'index':i}).start()
                    else:
                        threading.Thread(target=download_recursive,args=(gd,size,set_split_size,fileid,PickleFile),kwargs={'download':1,'index':i}).start()
                temp=[]
                start=time.time()
                while 1:
                    if len(temp)>=split_num_:break
                    time.sleep(25)
                    if time.time()-start>480:print("OUT!!!")
                    print(len(temp))
                    for j in range(0,len(gd.downloader)):
                        if j in temp:continue
                        if gd.downloader[j]:
                            
                            temp.append(j)
                fh=b''
                for jj in gd.downloader:
                    fh+=jj
                filepath=dst+'\\'+name_
                print('花費時間: ',time.time()-start)
                with io.open(filepath,'wb') as f:
                    
                    f.write(fh)
                
                return
            try:
                
                if 'application/vnd.google-apps' in str(mimeType_):
                    if 'spreadsheet' in str(mimeType_):
                        mimeType_='text/csv'
                        name_=name_+'.csv'
                    elif 'document'in str(mimeType_):
                        mimeType_='application/msword'
                        name_=name_+'.csv'
                    
                        request = gd.service.files().export_media(fileId=fileid,mimeType=mimeType_)
                else:
                    request = gd.service.files().get_media(fileId=fileid)
                print(name_,mimeType_)
                if not fh:
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
            except Exception as err:
                    print(err)
                    print(fileid,name_,mimeType_)
                    if 'malware' in str(err) or 'spam' in str(err):
                        request = gd.service.files().get_media(fileId=fileid,acknowledgeAbuse=True)
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
    
def Exapnd_dict_alphabet_number():
    count=26
    count1=1
    count2=1
    count3=0

    s3=0
    while count<1500:
        
        count+=1
        
        if count1%27==0:
            count1=1
            count2+=1
        count2=count2%27
        if count2%27==0:

            count2=1
            count3+=1
            s3=1
        count3=count3%27
        if count3%27==0 and s3:
            count3=1
        if count3:
            alp=num_alphabet[count3]+num_alphabet[count2]+num_alphabet[count1%27]
        else:
            alp=num_alphabet[count2]+num_alphabet[count1]
        num_alphabet[count]=alp
        alphabet_num[alp]=count
        count1+=1
Exapnd_dict_alphabet_number()

def RGB_to_HEX(R,G,B):

    Rh=hex(R).replace('0x','')
    Gh=hex(G).replace('0x','')
    Bh=hex(B).replace('0x','')
    return '#'+(Rh+Gh+Bh)
    
def HEX_to_RGB(Hex):
    if re.findall("#",str(Hex),re.IGNORECASE):
        index=Hex.find("#")
        Rh=Hex[index+1:index+3]
        Gh=Hex[index+3:index+5]
        Bh=Hex[index+5:index+7]
        
        R=int(Rh,16)
        G=int(Gh,16)
        B=int(Bh,16)
        
        return R,G,B


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
    size=int(size)
    for i in range(4,0,-1):
      if size>dict_[i]:
          return str(round(size/dict_[i],2))+' '+dict_unit[i]
      if i==1:
          return str(size)+' Byte'
class Drive:
    all_folder={}

    PickleFile=''
    json_path=''
    def __init__(self,*credential_path):
        
        self.downloader=[]
        self._download_path=Download_path
        if credential_path:
                        if '.json' in str(credential_path):
                                
                                self.chose_json(credential_path[0])
                        elif '.pickle' in str(credential_path):
                                self.chose_pickle(credential_path[0])
                        else:
                            raise TypeError(" Only acccept json or pickle file")
                            return 
                            
    @property
    def Download_path(self):
        return self._download_path
    @Download_path.setter
    def Download_path(self,setpath):
         if os.path.exists(setpath):
           self._download_path=setpath
         else:
             print("Path is not exist")
    def chose_json(self,path):
            
            self.json_path=path
            self.service,self.service_sheet=self.main()
            
            self.json_path=path
    def chose_pickle(self,path):
            
            
            self.PickleFile=path
            self.service,self.service_sheet=self.main()
    def emptytrash(self):
        self.service.files().emptyTrash().execute()
    def get_newest_file(self,**kwargs):
        now=datetime.datetime.now()
        if kwargs:
            if 'datetime' in kwargs:
                now=kwargs['datetime']
        
        today=datetime.datetime(now.year,now.month,now.day,0,0)
        timeutc_8=today-datetime.timedelta(hours=8)
        lasttime=timeutc_8.strftime("%Y-%m-%dT%H:%M:%S")
        x=self.service.files().list(                   q="createdTime>"+"'"+lasttime+"'"+"and trashed=False" ,
                                                  spaces='drive',
                                                  fields='*',
                                                  pageSize=1000,
                                                  pageToken=None,supportsAllDrives=True).execute()
        new_arr=[]
        month_n=now.month
        if now.month<10:
                month_n='0'+str(now.month)
        day_n=now.day
        if now.day<10:
                day_n='0'+str(now.day)
        todayfile=str(now.year)+str(month_n)+str(now.day)+'.csv'
        dict_temp={}
        if os.path.isfile(todayfile):
            with open(todayfile,'r') as tf:
                for i in tf:
                        
                    id_,name=i.split(',')
                    dict_temp[id_]=name
        for i in x['files']:
            fileid=i['id']
            if fileid  in dict_temp:continue
            createdTime=i['createdTime']
            year=int(createdTime[:4])
            month=int(createdTime[5:7])
            day=int(createdTime[8:10])
            hour=int(createdTime[11:13])
            mins=int(createdTime[14:16])
            secs=int(createdTime[17:19])
            newtime=datetime.datetime(year,month,day,hour,mins,secs)+datetime.timedelta(hours=8)
            link=i['webViewLink']
            
            name=i['name']
            if fileid not in dict_temp:
                new_arr.append(i)
                dict_temp[fileid]=name
                self.change_permissions(fileid,role='reader')
        if new_arr:
            
            with open(todayfile,'a+') as tf:
                for i in new_arr:
                    name=i['name']
                    fileid=i['id']
                    tf.write(str(fileid))
                    tf.write(',')
                    tf.write(name)
                    tf.write('\n')
            return  new_arr
    def find_folder(self):
                results = self.service.files().list(q="mimeType='application/vnd.google-apps.folder' and trashed=false",
                                              spaces='drive',
                                              fields='nextPageToken,files(ownedByMe,mimeType,id, name)',
                                              
                                              pageToken=None).execute()
                
                return results
    def find_file(self,filename,*args,**kwargs):
          show_=1
          if args:
            show_=0
    
          items=self.list_all_file(**kwargs,view=0)          
                    
          
          if not items:return 0
          itemfind=[]
          similar=[]
          dict_of_all={}
          for __,item in items.items():
                
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
              filedict=self.service.files().get(fileId=i).execute()
              name_=filedict['name']
              mimetype=filedict['mimeType']
              dict_of_all[i]=name_,mimetype
              if show_:print('ID: '+i+" | Name: {:<20}".format(name_))
          if similar :
           if show_:print("找到相似名稱的項目:")
           for i in similar:
              filedict=self.service.files().get(fileId=i).execute()
              name_=filedict['name']
              mimetype=filedict['mimeType']
              dict_of_all[i]=name_,mimetype
              if show_:print('ID: '+i+" | Name: {:<20}".format(name_))
          if not itemfind and not similar:
              print("找不到項目")
              return 0
          list_=sorted(dict_of_all.items(),key=lambda b:b[1],reverse=True)
          
          dict_of_all={}
          for i in list_:
              key,filename_filetype=i
              filename,filetype=filename_filetype
              dict_of_all[key]=filename,filetype
          
          
          
          return dict_of_all
    def listdir(self,fileid):
        
        response=self.service.files().list(q="'"+str(fileid)+"' in parents",
                                          fields='files(id, name,mimeType,trashed)',
                                          pageSize=101
                                         ).execute()
        result_=response['files']
        if len(result_)>=100:
                dict_all=self.list_all_file(view=0,stop=10000)
                temp_dict=[]
                
           
                for i in dict_all:
                        
                        if dict_all[i]['parents'][0].lower().strip()==fileid.lower():
                                temp_dict.append(dict_all[i])
                           
               
                return temp_dict
        else:
                return result_
    
    def _requests_download(self,fileid,dst,**kwargs):
        filedict=self.service.files().get(fileId=fileid,fields='*').execute()
        name_=filedict['name']
        headers = {"Authorization": "Bearer A0AfH6SMCnI8nJDQ97yOKibKRx1mxvYQJV5RiduXB0cA1v3JPupkR780xgvS0FaIuvNUlY8EAuYy_HTTg78hLzrHifkEDlNkXCDFrYyj7oMbbwdWzijAqefm3UeFsIlN8tT37SXIvnmI51C1DGZfpju7cM4MdkBgsgiv2PSuJyhIWmq"}

        request1 = gd.service.files().get_media(fileId=fileid)
       # headers=request1.headers
        url="https://www.googleapis.com/drive/v3/files/" + fileid + "?alt=media"
        res = requests.get(url,headers=headers)
        print(res.status_code)
        filepath=dst+'\\'+name_
        start=0
        if kwargs:
                if 'st' in kwargs:
                        start=kwargs['st']
                        
        with open(filepath,'wb') as f:
                   f.write(res.content)
      
    def download_id(self,fileid,dst,*args,**kwargs):
        exception=[]
        size=0
        if kwargs:
                
                if 'exception' in kwargs:
                        exception=kwargs['exception']
                        
               
        if not args:
                self.count=0
                self.time_zero=time.time()
        else:
                self.count=args[0]
        if fileid:
            
            
            filedict=self.service.files().get(fileId=fileid,fields='*').execute()
            name_=filedict['name']
            
            if 'trashed' in filedict:
                trash_=filedict['trashed']
                if 'true' in str(trash_).lower().strip():return
            mimeType_=filedict['mimeType']
            ownedByMe_=filedict['ownedByMe']
            
            if 'false' in str(ownedByMe_):return
            if not args:
                print("Download " +name_ )
                
            if 'folder'in str(mimeType_):
                list_of_files=self.listdir(fileid)
                fold_name=dst+'\\'+str(name_)
               
                if not os.path.isdir(fold_name):
                    os.mkdir(fold_name)
                
               
                for each in list_of_files:
                    
                    self.download_id(each['id'],fold_name,1,exception=exception)
                   
                    
                print("Download 100%")
                
                return
            if 1:
                self.count+=1
                filepath=dst+'\\'+name_
                if 'size' in filedict:
                    size=filedict['size']
                    if int(size)>1024*1024*1024:
                          
                             filepath=dst+'\\[DriverBig]'+name_
                             with open(filepath,'wb') as f:
                                 
                                    f.write(b"")
                             return
                            
              
                
               
                if 'application/vnd.google-apps' in str(mimeType_):
                    
                    if 'spreadsheet' in str(mimeType_):
                        mimeType_='text/csv'
                        name_=name_+'.csv'
                  
                    elif 'document'in str(mimeType_) and 'doc' in str(name_).lower():
                        mimeType_='application/msword'
                        name_=name_+'.docx'
                    elif 'document'in str(mimeType_):
                        mimeType_='text/plain'
                        name_=name_+'.txt'
                    request = self.service.files().export_media(fileId=fileid,mimeType=mimeType_)
                else:
                   
                    
                    request = self.service.files().get_media(fileId=fileid)
                fh = io.BytesIO()
                
                downloader = MediaIoBaseDownload(fh, request)
                if fileid in exception:return
                try:
                        done = 0
                                        # 如果想下載包含病毒的檔案 ,acknowledgeAbuse=True
                        while not done :# request = gd.service.files().get_media(fileId=fileid,acknowledgeAbuse=True)
                            status, done = downloader.next_chunk()
                            if not args:
                              print("Download %d%%" % int(status.progress() * 100))
                        
                        filepath=dst+'\\'+name_
                        with io.open(filepath,'wb') as f:
                            fh.seek(0)
                            f.write(fh.read())
                except:
                        filepath=dst+'\\[DriverError]'+name_
                        with io.open(filepath,'wb') as f:
                            fh.seek(0)
                            f.write(fh.read())
            
    def download(self,file,*dst):
        if dst:
            dst=dst[0]    
        file=str(file)
        if not dst:
            print("Download to default path: "+self._download_path)
           
            dst=self._download_path
        
        dict_of_find=self.find_file(file,1,view=0)#view= display found file
        temp_dict={}
        fileid=0
        count=0
        alreadyprint=[]
        
        if dict_of_find:
                        print("File was found:") 
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
                                  
                                        
                                        size=self.get_size(each['id'])
                                        print('╘═'+format_str(3,str(count)+'.'),format_str(40,each['name']),'| Size:',size)
                                        alreadyprint.append(each['id'])
                                        temp_dict[str(count)]=each['id'],each['name']
                                        
                                        count+=1
                            else:
                                size=self.get_size(i)
                                print(format_str(3,str(count)+'.'),format_str(40,name_),'| Size:',size)
                            
                                temp_dict[str(count)]=i,name_
        if not temp_dict:return
        print("\nIf wnat to download:  1. %s  >> press 1 "%(temp_dict[str(1)][1]),'(* press enter to skip ) ')
        if count>1:
          chose=input('[* Choose more than one : press 1~3(choose: 1,2,3) or 1,3,5(choose: 1,3,5)]\n')
        else:
          chose=input('>>\n')            
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
            
            filedict=self.service.files().get(fileId=fileid).execute()
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
                  
                    elif 'document'in str(mimeType_) and 'doc' in str(name_).lower():
                        mimeType_='application/msword'
                        name_=name_+'.docx'
                    elif 'document'in str(mimeType_):
                        mimeType_='text/plain'
                        name_=name_+'.txt'
                    request = self.service.files().export_media(fileId=fileid,mimeType=mimeType_)
            else:
                    request = self.service.files().get_media(fileId=fileid)
            fh = io.BytesIO()
            downloader = MediaIoBaseDownload(fh, request)
            try:
                        
                        done = 0
                                        # 如果想下載包含病毒的檔案 ,acknowledgeAbuse=True
                        while not done :# request = gd.service.files().get_media(fileId=fileid,acknowledgeAbuse=True)
                            status, done = downloader.next_chunk()
                            print("Download %d%%" % int(status.progress() * 100))
                       
                        filepath=dst+'\\'+name_
                        with io.open(filepath,'wb') as f:
                            fh.seek(0)
                            f.write(fh.read())
            except Exception as er:
                        print(er)
                       
                        filepath=dst+'\\[Error]'+name_
                        with io.open(filepath,'wb') as f:
                            fh.seek(0)
                            f.write(fh.read())
            
            
    def change_permissions(self,file_id,**kwarg):# 將某ID的檔案權限更改
        type_='anyone'
        role_='writer'
        if kwarg:
            if 'type' in kwarg:
                type_=kwarg['type']
            if 'role' in kwarg:
                role_=kwarg['role']
        self.service.permissions().create(fileId=file_id,
                                             body= {
                                            'role': role_,
                                            'type': type_,
                                              }
                                            ).execute()
    
    def create_newsheet(self,*filename,**folder):
            folderid=None
            folderid2=None
            foldername=''
            newfolder=None
            if not filename:
                filename='NewSheet'+str(timenow)
            
            if folder :
                if 'newfolder'in folder:
                         newfolder=folder['newfolder']
                if 'foldername' in folder:
                        foldername=folder['foldername'] # second folder
                if 'folder_id' in folder:
                        folderid=folder['folder_id'] #top folder
                if newfolder:
                       if not folderid2:
                          folderid2=self.mkdir(newfolder,folder_id=folderid) 
                
                if foldername: 
                  folderid2=self.find_folder_id(foldername)
                  if not folderid2:
                          folderid2=self.mkdir(foldername,folder_id=folderid)
                if folderid2 and folderid:
                    file_metadata = {
                    
                    'name': filename,
                    'parents':[folderid2],
                    'mimeType': 'application/vnd.google-apps.spreadsheet'
                    }
                    file = self.service.files().create(body=file_metadata,supportsAllDrives=True).execute()
                   
                    
                    return file['id']
                
                else:
                    folderid=self.mkdir(foldername)
                    file_metadata = {
                    'name': filename,
                    'parents':[folderid],
                    'mimeType': 'application/vnd.google-apps.spreadsheet'
                    }
                    file = self.service.files().create(body=file_metadata,supportsAllDrives=True).execute()
                    return file['id']
                
            file_metadata = {
            'name': filename,
            'mimeType': 'application/vnd.google-apps.spreadsheet'
            }
            
            file = self.service.files().create(body=file_metadata,supportsAllDrives=True).execute()
            
            return file['id']
    def create_doc(self,*filename):
        if not filename:
            filename="NewDoc"
        file_metadata = {
        'name': filename,
        'mimeType': 'application/vnd.google-apps.document'
        }
        
        file = self.service.files().create(body=file_metadata,supportsAllDrives=True).execute()
        
        return file['id']
    def create_Slides(self,*filename):
        if not filename:
            filename="NewSlide"
        file_metadata = {
        'name': filename,
        'mimeType': 'application/vnd.google-apps.presentation'
        }
        
        file = self.service.files().create(body=file_metadata,supportsAllDrives=True).execute()
        
        return file['id']
    def create_form(self,*filename):
        if not filename:
            filename="NewSlide"
        file_metadata = {
        'name': filename,
        'mimeType': 'application/vnd.google-apps.form'
        }
        
        file = self.service.files().create(body=file_metadata,supportsAllDrives=True).execute()
        
        return file['id']
    def find_folder_id(self,foldername):
          results = self.service.files().list(     q="mimeType='application/vnd.google-apps.folder' and trashed=false and name='"+foldername+"'",
                                              spaces='drive',                                              
                                              fields='nextPageToken, files(id, name)',
                                              pageToken=None,supportsAllDrives=True).execute()
          items = results.get('files', [])
          
          if not items:return 0
          
          for item in items:
                 id_=item['id']
                 name_=item['name']
                 
                 if foldername==name_:
                     #folderID=id_
                     return id_
          return 0
    def upload(self,filepath,*dstpath,**kwarg):# 上傳檔案到特定資料夾
        filename=''
        foldername=''
        initial_foldername=''
        initial_id=''
        parents_id=''
        folderID=''
        folder_id=''
        if kwarg:
                if 'folder_id' in kwarg:# 特定資料夾的ID
                         folder_id=kwarg['folder_id']
                if 'initial_foldername' in kwarg:
                        initial_foldername=kwarg['initial_foldername']
                if 'parents_id' in kwarg:# 特定資料夾的ID 等同 folder_id
                        parents_id=kwarg['parents_id']
                if 'initial_id' in kwarg:
                        initial_id=kwarg['initial_id']
        if parents_id or folder_id:
                    folderID=parents_id
        if os.path.isdir(filepath) :# dir file
            filepath=os.path.abspath(filepath)
            foldername=os.path.basename(filepath)
            
            
            if dstpath:
                
                foldername=os.path.dirname(dstpath[0])
           
                if not foldername:
                    foldername=dstpath[0]
                folderID=self.find_folder_id(foldername)
                if parents_id:
                    folderID=parents_id
                if folderID:
                    file_metadata = {
                    'name': foldername,
                    'parents' : [folderID],
                    'mimeType': 'application/vnd.google-apps.folder'
                    }
                    file = self.service.files().create(body=file_metadata).execute()
                else:
                     
                     file_metadata = {
                    'name': foldername,
                   
                    'mimeType': 'application/vnd.google-apps.folder'
                     }

                     file = self.service.files().create(body=file_metadata).execute()
                     folderID=file['id']
                     
            else:
                    if not folderID:
                        
                        
                        folderID=self.find_folder_id(foldername)
                      
                        
                    if not folderID:
                            file_metadata = {
                                'name': foldername,
                                
                                'mimeType': 'application/vnd.google-apps.folder'
                                }
                            file = self.service.files().create(body=file_metadata).execute()
                            
                    else:
                            file_metadata = {
                            'name': foldername,
                            'parents' : [folderID],
                            'mimeType': 'application/vnd.google-apps.folder'
                            }
                            file = self.service.files().create(body=file_metadata).execute()
                    
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
            print("Mimetype not found , try to upload this folder directly")
            foldername=os.path.basename(dstpath[0])
            filename=os.path.basename(filepath)
            Extension=os.path.splitext(filename)[-1].replace('.','')
            
            if Extension in mimetype_dict:
                
                file_type=mimetype_dict[Extension]
            else:file_type='application/octet-stream'
            

        
        
        
       
        if foldername:
          
          if not folderID and not parents_id:
              folderID=self.find_folder_id(foldername)
          
         
          if  not folderID and not parents_id:
              file_metadata = { 'name': filename}
          else:
              file_metadata = {
                'name': filename,
                'parents': [folderID]
                }       
        else:
              file_metadata = { 'name': filename}          
        
        if parents_id:
                    folderID=parents_id
                    file_metadata = {
                        'name': filename,
                        'parents': [folderID]
                        }
      
        media = MediaFileUpload(filepath,
                                mimetype=file_type,
                                resumable=True)
            
        file = self.service.files().create(body=file_metadata,
                                            media_body=media,supportsAllDrives=True,
                                            fields='id').execute()
        
        print("Upload Finish !")
        return file
    def delete_id(self,fileid):
        file_=self.service.files().get(fileId=fileid).execute()
        print("Delete: ",file_['name'])
        try:
            self.service.files().delete(fileId=fileid).execute()
        except :
          print("Delete Error")
    def delete(self,filename):
      filename=str(filename)
      dict_of_find=self.find_file(filename,1,view=0)#view= display found file
      temp_dict={}
      fileid=0
      count=0
      alreadyprint=[]
      if dict_of_find:
                    print("\n=== === Files List === ===")
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
      
      print("\nIf wnat to delete:  1. %s  >> press 1"%(temp_dict[str(1)][1]),' (* press enter to skip )')
      if count>1:
        chose=input('[* Choose more than one : press 1~3(choose: 1,2,3) or 1,3,5(choose: 1,3,5)]\n')
      else:
            chose=input('>>\n')
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
      file_=self.service.files().get(fileId=fileid).execute()
      print("Delete: ",file_['name'])
      try:
            self.service.files().delete(fileId=fileid).execute()
      except :
          print("Delete Error")
    def get_weblink(self,fileid):
        # fields = get what you want
        # file resource url=https://developers.google.com/drive/api/v3/reference/files
        results = self.service.files().get(fileId=fileid,fields='webViewLink').execute()['webViewLink']
        return results
    def get_storageQuota(self):
       storage=self.service.about().get(fields='*').execute()['storageQuota']
    
       alreadyuse='Usage: '+size_byte(int(storage['usage']))
       if 'limit' in storage:
         limit='Limit Usage: '+size_byte(int(storage['limit']))
       trash="Trash Usage: " +size_byte(int(storage['usageInDriveTrash']))
       
       return alreadyuse,limit
    def get_size(self,fileid,*arg):
      
        dict_={1:1024,2:1024**2,3:1024**3,4:10**12}
        dict_unit={1:'KB',2:'MB',3:'GB',4:'TB'}
        try:
          type_=self.get_type(fileid)
        except:type_=''
        if 'folder' in str(type_).lower():
            sizes=0
            list_files=self.listdir(fileid)
            for j in list_files:
             
               
                if 'true' in str(j['trashed']).lower():
                
                        continue
                size=self.get_size(j['id'],0)
                sizes+=int(size)
            if arg:  return sizes
            
            return size_byte(sizes)
        
        
        results = self.service.files().get(fileId=fileid,fields='size',supportsAllDrives=True).execute()['size']
        
        if arg:  return results

        return size_byte(int(results))
    def get_type(self,fileid):
        # fields = get what you want
        # file resource url=https://developers.google.com/drive/api/v3/reference/files
        results = self.service.files().get(fileId=fileid,fields='mimeType').execute()['mimeType']
        return results
    def get_lastModifyingUser(self,fileid):
        # fields = get what you want
        # file resource url=https://developers.google.com/drive/api/v3/reference/files
        results = self.service.files().get(fileId=fileid,fields='lastModifyingUser').execute()['lastModifyingUser']
        return results
    def get_shared(self,fileid):
        # fields = get what you want
        # file resource url=https://developers.google.com/drive/api/v3/reference/files
        results = self.service.files().get(fileId=fileid,fields='shared').execute()['shared']
        return results
    def main(self):
        """Shows basic usage of the Drive v3 API.
        Prints the names and ids of the first 10 files the user has access to.
        """
        creds = None
        service_account=0
        # The file token.pickle stores the user's access and refresh tokens, and is
        # created automatically when the authorization flow completes for the first
        # time.
        
        if os.path.exists(self.PickleFile):
            
            with open(self.PickleFile, 'rb') as token:
                creds = pickle.load(token)
                
       
        # If there are no (valid) credentials available, let the user log in.
        try:
         if not creds or not creds.valid:
            
            if creds and creds.expired and creds.refresh_token:
                creds.refresh(Request())
            else:
                if os.path.exists(self.json_path):
                 with open(self.json_path,'r') as jsooon:
                    if 'service_account' in str(jsooon.read()):
                        service_account=1
                if service_account:
                    creds = ServiceAccountCredentials.from_json_keyfile_name(
                        self.json_path, scopes=SCOPES)
                    
                    
                else:
                    flow = InstalledAppFlow.from_client_secrets_file(
                        self.json_path, SCOPES)
                    creds = flow.run_local_server(port=8012)
                    print(creds)
                    # Save the credentials for the next run
                    self.PickleFile='token.pickle'
                    with open('token.pickle', 'wb') as token:
                
                        pickle.dump(creds, token)
        except Exception as Errorlogin:
            print(Errorlogin)
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
        
        
        request = self.service_sheet.spreadsheets().values().update(spreadsheetId=sheetid, range=workRange,valueInputOption='RAW',body=writeData )
        response = request.execute()   
    def sheet_append(self,sheetid,workRange,content):
        if 'list' not in str(type(content)):
                content=[content]
        writeData = {}
        
     
        writeData['values']= content
        writeData['majorDimension']="ROWS"
        writeData['range']= workRange
        
       
        
        request = self.service_sheet.spreadsheets().values().append(spreadsheetId=sheetid, range=workRange,valueInputOption='RAW', insertDataOption='INSERT_ROWS', body=writeData )
        response = request.execute()
    def get_initial_folderid(self,*arg):
        allid=self.find_folder()['files']
        temp=[]
        if arg:
            temp=arg[0]
        for i in allid:
            temp.append(i['id'])
        dict_=self.list_all_file(view=0,stop=100)

        for __,item in dict_.items():
                    
                     id_=item['id']
                     name_=item['name']
                     if 'ownedByMe' in item:
                             ownedByMe=item['ownedByMe']
                             if 'true' not in str(ownedByMe).lower():continue
                     type_=item['mimeType']
                     if 'parents' not in item:continue
                     parents_=item['parents'][0]
                     trashed_=item['trashed']
                     if 'true' in str(trashed_).lower():continue
                     if parents_ not in temp:
                        
                         break
        ini_=self.service.files().get(fileId=parents_,fields='*'). execute()
        
        ini_id=ini_['id']
        ini_name_=ini_['name']
        if ('my' not in ini_name_ and '我的雲端硬碟' not in ini_name_) or 'parents'  in ini_:self.get_initial_folderid(temp)
        
        return ini_,dict_,ini_id
    def get_filelist(self):
        print("Wait...")     
        dict_=self.list_all_file(view=0)
        if dict_:
          sizeall=0
          print("Writing down...")
          print('Download Path: ',Download_path+'\\File_list.csv')
          with open (Download_path+'\\File_list.csv','w+',encoding='utf-8-sig')as store_csv:
               store_csv.write('Name')
               store_csv.write(',')
               store_csv.write('ID')
               store_csv.write(',')
               store_csv.write('MimeType')
               store_csv.write(',')
               store_csv.write('Size')
               store_csv.write(',')
               store_csv.write('Link')
               store_csv.write(',')
               store_csv.write('LastModifyUser')
               store_csv.write(',')
               store_csv.write('Shared')

               store_csv.write(',')
               store_csv.write('Createtime')
               store_csv.write('\n')
               for __,item in dict_.items():
                     
                     id_=item['id']
                     
                     name_=item['name']
                     type_=item['mimeType']
                     if 'size' not in item:
                         size=0
                     else:size=item['size']
                     
                     sizeall+=int(size)
                     link=item['webViewLink']
                     shared=item['shared']
                     if not shared:
                         shared="Not shared"
                     else:
                         shared="Shared"
                     user=item['lastModifyingUser']['emailAddress']
                     sizef=lambda s:'-' if not s else size_byte(s)
                     createtime=''
                     if 'createdTime' in item:
                       createtime=item['createdTime']
                     store_csv.write(str(name_))
                     store_csv.write(',')
                     store_csv.write(str(id_))
                     store_csv.write(',')
                     store_csv.write(str(type_))
                     store_csv.write(',')
                     store_csv.write(sizef(size))
                     store_csv.write(',')
                     store_csv.write(str(link))
                     store_csv.write(',')
                     store_csv.write(str(user))                 
                     store_csv.write(',')
                     store_csv.write(shared)
                     store_csv.write(',')
                     store_csv.write(str(createtime))
                     store_csv.write('\n')
               store_csv.write(str('總計'))
              
               store_csv.write(',')
               store_csv.write(str(','))
               store_csv.write(',')
               store_csv.write(sizef(sizeall))
               store_csv.write(',')
               store_csv.write(str(','))
               store_csv.write(',')
               store_csv.write(',')                 
               store_csv.write(',')
               store_csv.write(str(','))
               store_csv.write(str(','))
               store_csv.write('\n')
               print('OK')
    def list_all_file(self,*args,**kwargs):
         view=1
         dict_all={}
         stopnum=0
         if kwargs:
               
             if'view' in kwargs:
                 if kwargs['view']==0:
                     view=0
             if 'dict_all' in (kwargs):
                 dict_all=kwargs['dict_all']
             if 'stop' in kwargs:
                     stopnum=kwargs['stop']
         if not args:
             count=0
             pageToken=None
             dict_all={}
             
         else:
           
             
             pageToken=args[0]
             count=args[1]
         results = self.service.files().list(
                                              spaces='drive',
                                              fields='*',
                                              pageSize=1000,
                                              pageToken=pageToken).execute()
         
         if stopnum:# 避免過久
                 if len(dict_all) > int(stopnum):return dict_all
         for j in results['files']:
             count+=1
             if j['ownedByMe']:
               
               if view:print(str(count)+'. '+str(j['name']),' - OwnedByme')
             else:
               
               if view:print(str(count)+'. '+str(j['name']),' - OwnedByOther')
             dict_all[j['id']]=j
                 
         if not results['files'] or 'nextPageToken' not in results:
             
             return dict_all
         
         return self.list_all_file(results['nextPageToken'],count,dict_all=dict_all,view=view,stop=stopnum)
    def mkdir(self,foldername,folder_id=None):
                if folder_id:
                        file_metadata = {
                            'name': foldername,
                            'parents':[folder_id],
                    
                            'mimeType': 'application/vnd.google-apps.folder'
                            }
                        file = self.service.files().create(body=file_metadata,supportsAllDrives=True).execute()
                else:
                        
                        file_metadata = {
                            'name': foldername,
                            
                            'mimeType': 'application/vnd.google-apps.folder'
                            }
                        file = self.service.files().create(body=file_metadata,supportsAllDrives=True).execute()
                return file['id']
    def move(self,fileId,*folder,**kwargs):
            if kwargs:
                    if 'folder_id' in kwargs:
                            folder_id=kwargs['folder_id']
            if folder:
              folder_id=self.find_folder_id(folder)
            if not folder_id:print("Destination of folder was not found")
            if folder_id:
               file = self.service.files().get(fileId=fileId,
                                         fields='parents').execute()
               previous_parents = ",".join(file.get('parents')) 
               file = self.service.files().update(fileId=fileId,
                                            addParents=folder_id,
                                            removeParents=previous_parents,
                                            fields='id, parents').execute()
    def delete_emptyfolder(self):
        result=self.find_folder()
        emptyfolder={}
        for each in result['files']:
          own=each['ownedByMe']
          if own:
              
              list_files=self.listdir(each['id'])
              if not list_files:
                  emptyfolder[each['id']]=each
        count=1
        if emptyfolder:
            temp_dict={}
            print("Find empty folder:")
        for j in emptyfolder:
            name=emptyfolder[j]['name']
            
            print(format_str(3,str(count)+'.'),format_str(40,name))
            temp_dict[str(count)]=j
            count+=1
        chose=input("If yout want to delete , Press 'Delete' :\n")
        if 'delete' in str(chose).lower():
            chose=input("Chose which one to delete , if all Press 'all' :\n")
            if str(chose) in temp_dict:
                        fileid=temp_dict[str(chose)]
            elif chose:
                if '~' in str(chose):
                    file_=chose.split('~')
                    for chose_ in range(int(file_[0]),int(file_[-1])+1):
                        if str(chose_) in temp_dict:
                            fileid=temp_dict[str(chose_)]
                            
                            self.delete_id(fileid)
                         
                    return    
                elif ',' in str(chose):
                    file_=chose.split(',')
                    for chose_ in file_:
                        if str(chose_) in temp_dict:
                            fileid,name_s=temp_dict[str(chose_)]
                            self.delete_id(fileid)
                    return

class Sheet:
    formula_use_data='for_formula_use_!A1'
    def __init__(self,*sheetid,**kwargs):
        
        
        
        if sheetid:
           self._id=sheetid[0]
           
           try:
             self.getsheet()
           except Exception as Err:
                   if 'not found' in str(Err):
                       print("Your SheetID was not found")
                   else:
                       raise TypeError("Login drive first >> drive = gdi.Drive(json or pickle path)")
                   
                   
        else:
            timenow=datetime.datetime.now().strftime("%H%M%S")
            if kwargs:
                if 'folder' in kwargs:
                                foldername=kwargs['folder']
                for i in kwargs:
                  find_name=re.findall("page|new_?page|new_?name|new_?title|n\S?ame|title|nam_?|",str(i),re.IGNORECASE)
                  if find_name and len(find_name)<=2:
                   
                    new_name=kwargs[find_name[0]]
                    self._id=self.create(str(new_name),folder=foldername)
            else:
              self._id=self.create("NewSheet"+str(timenow))
              
              
              print("Can't find SheetID , automake sheetname: NewSheet"+str(timenow))
              print("Spreadsheet ID: " ,self._id)
    def merge(self,range_,**sheetid):
         
         list_=self.__range__return(range_)
         col1,row1=list_[0]
         col2,row2=list_[-1]
      
         sheetid=0
         if sheetid:
             sheetid=sheetid['sheetid']
         
         top_header_format = [
            {'mergeCells': {
                'mergeType': 'MERGE_ALL',
                'range': {
                    'endColumnIndex': col2,
                    'endRowIndex': row2,
                    'sheetId': sheetid,
                    'startColumnIndex': col1-1,
                    'startRowIndex':row1-1
                }
            }}
            ]
         self.service_sheet.spreadsheets().batchUpdate(
            spreadsheetId=self._id ,
            body={'requests': top_header_format}
        ).execute()
    def ___get__index_start_end(self,Colrange):
                
                if ',' in Colrange:
                    Colrange_1,Colrange_2=Colrange.split(',')
                    if not Colrange_1.isnumeric() and not Colrange_2.isnumeric():
                       
                        Colrange_1,Colrange_2=(self._find_columb_alphabet_to_number(Colrange))
                    elif not Colrange_1.isnumeric() and  Colrange_2.isnumeric():
                        Colrange_1,_=self._find_columb_alphabet_to_number(Colrange_1)
                    elif Colrange_1.isnumeric() and  not Colrange_2.isnumeric():
                        Colrange_2,_=self._find_columb_alphabet_to_number(Colrange_2)
                    
                elif Colrange and ":" in str(Colrange):
                    Colrange_1,Colrange_2=Colrange.split(':')
                    
                    if not Colrange_1.isnumeric() and not Colrange_2.isnumeric():
                             Colrange_1,Colrange_2=self._find_columb_alphabet_to_number(Colrange)
                    elif not Colrange_1.isnumeric() and  Colrange_2.isnumeric():
                        Colrange_1,_=self._find_columb_alphabet_to_number(Colrange_1)
                    elif Colrange_1.isnumeric() and  not Colrange_2.isnumeric():
                        Colrange_2,_=self._find_columb_alphabet_to_number(Colrange_2)
                else:
                    if Colrange.isnumeric():
                        Colrange_1,Colrange_2=int(Colrange),int(Colrange)
                    else:
                        Colrange_1,Colrange_2=(self._find_columb_alphabet_to_number(Colrange))
                        if Colrange_2==None:
                            Colrange_2=Colrange_1
                return Colrange_1,Colrange_2
    def adjust_col_row(self,Colrange,Col_pixel,Rowrange,Row_pixel,*sheetID,**kwargs):
               
                Colrange=str(Colrange)
                Rowrange=str(Rowrange)
                
                sheetId=  self.__arg_return_sheetid(*sheetID,**kwargs)
                
                if sheetId==None:return "Not found this sheet name"
                sheetname=self.__arg_return_sheetname(sheetId)
                
                Rowrange_1,Rowrange_2=1,1
                Colrange_1,Colrange_2=1,1
                if 'all' in Colrange:
                    col,row=self.get_sheet_size(sheetname=sheetname)
                    Colrange_1=1
                    Colrange_2=col
                    
                else:
                    Colrange_1,Colrange_2=self.___get__index_start_end(Colrange)
                  
                if 'all' in Rowrange:
                    col,row=self.get_sheet_size(sheetname=sheetname)
                    Rowrange_1=1
                    Rowrange_2=row
                  
                elif Rowrange and ":" in str(Rowrange):
                    Rowrange_1,Rowrange_2=Rowrange.split(":")
                    if not Rowrange_1:
                        Rowrange_1=0
                elif  Rowrange and "," in str(Rowrange):
                    Rowrange_1,Rowrange_2=Rowrange.split(":")
                    if not Rowrange_1:
                        Rowrange_1=0
                else:
                    Rowrange_1,Rowrange_2=int(Rowrange),int(Rowrange)
                
               
                checknum=lambda x : 1 if int(x)<1 else int(x)
                Colrange_1=checknum(Colrange_1)
                Colrange_2=checknum(Colrange_2)
                Rowrange_1=checknum(Rowrange_1)
                Rowrange_2=checknum(Rowrange_2)
                body={
                      "requests": [
                        {
                          "updateDimensionProperties": {
                            "range": {
                              "sheetId": sheetId,
                              "dimension": "COLUMNS",
                              "startIndex": int(Colrange_1)-1,
                              "endIndex":   int(Colrange_2)
                            },
                            "properties": {
                              "pixelSize": Col_pixel
                            },
                            "fields": "pixelSize"
                          }
                        },
                        {
                          "updateDimensionProperties": {
                            "range": {
                              "sheetId": sheetId,
                              "dimension": "ROWS",
                              "startIndex": int(Rowrange_1)-1,
                              "endIndex":   int(Rowrange_2)
                            },
                            "properties": {
                              "pixelSize": Row_pixel
                            },
                            "fields": "pixelSize"
                          }
                        }
                      ]
                    }
                request = self.service_sheet.spreadsheets().batchUpdate(spreadsheetId=self._id ,body=body).execute()
    def adjust_row(self,Rowrange,Row_pixel,*sheetID,**kwargs):
       
                Rowrange=str(Rowrange)
                sheetId=  self.__arg_return_sheetid(*sheetID,**kwargs)
                if sheetId==None:return "Not found this sheet name"
                
                sheetname=self.__arg_return_sheetname(sheetId)
                if 'all' in Rowrange:
                    col,row=self.get_sheet_size(sheetname=sheetname)
                    Rowrange_1=1
                    Rowrange_2=row
                  
                elif Rowrange and ":" in str(Rowrange):
                    Rowrange_1,Rowrange_2=Rowrange.split(":")
                    if not Rowrange_1:
                        Rowrange_1=0
                elif  Rowrange and "," in str(Rowrange):
                    Rowrange_1,Rowrange_2=Rowrange.split(":")
                    if not Rowrange_1:
                        Rowrange_1=0
                else:
                    Rowrange_1,Rowrange_2=int(Rowrange),int(Rowrange)
                
               
                checknum=lambda x : 1 if int(x)<1 else int(x)
                
                Rowrange_1=checknum(Rowrange_1)
                Rowrange_2=checknum(Rowrange_2)
                    
                
                if not (Rowrange_1): Rowrange_1=1
                body={
                      "requests": [
                      
                        {
                          "updateDimensionProperties": {
                            "range": {
                              "sheetId": sheetId,
                              "dimension": "ROWS",
                              "startIndex": int(Rowrange_1)-1,
                              "endIndex":   int(Rowrange_2)
                            },
                            "properties": {
                              "pixelSize": Row_pixel
                            },
                            "fields": "pixelSize"
                          }
                        }
                      ]
                    }
                request = self.service_sheet.spreadsheets().batchUpdate(spreadsheetId=self._id ,body=body).execute()
                
    def adjust_col(self,Colrange,Col_pixel,*sheetID,**kwargs):
                Colrange=str(Colrange)
                
                sheetId=  self.__arg_return_sheetid(*sheetID,**kwargs)
                if sheetId==None:return "Not found this sheet name"
                sheetname=self.__arg_return_sheetname(sheetId)
                Colrange_1,Colrange_2=1,1
                if 'all' in Colrange:
                    
                    col,row=self.get_sheet_size(sheetname=sheetname)
                    Colrange_1=1
                    Colrange_2=col
                    
                else:
                    Colrange_1,Colrange_2=self.___get__index_start_end(Colrange)
                  

                checknum=lambda x : 1 if int(x)<1 else int(x)
                Colrange_1=checknum(Colrange_1)
                Colrange_2=checknum(Colrange_2)
               
                body={
                      "requests": [
                        {
                          "updateDimensionProperties": {
                            "range": {
                              "sheetId": sheetId,
                              "dimension": "COLUMNS",
                              "startIndex": int(Colrange_1)-1,
                              "endIndex":   int(Colrange_2)
                            },
                            "properties": {
                              "pixelSize": Col_pixel
                            },
                            "fields": "pixelSize"
                          }
                        }
                      ]
                    }
                request = self.service_sheet.spreadsheets().batchUpdate(spreadsheetId=self._id ,body=body).execute()
    
    def delete_row(self,range_,*sheetID,**kwargs):
               start,end=range_
               sheetId=  self.__arg_return_sheetid(*sheetID,**kwargs)
               if sheetId==None:return "Not found this sheet name"
               if start<1:start=1
               body=   {
                  "requests": [
                    {
                      "deleteDimension": {
                        "range":{
                             "sheetId": sheetId,
                              "dimension": "ROWS",
                              "startIndex": int(start)-1,
                              "endIndex": end


                            }}}]}
                       
               request = self.service_sheet.spreadsheets().batchUpdate(spreadsheetId=self._id ,body=body).execute()
          
    def append_row(self,length,*sheetID,**kwargs):
               sheetId=  self.__arg_return_sheetid(*sheetID,**kwargs)
               if sheetId==None:return "Not found this sheet name"
               
               body=   {
                  "requests": [
                    {
                      "appendDimension": {
                        "sheetId": sheetId,
                        "dimension": "ROWS",
                        "length": length
                      }
                    }
                    
                ]}
               request = self.service_sheet.spreadsheets().batchUpdate(spreadsheetId=self._id ,body=body).execute()
    def _find_columb_alphabet_to_number(self,text):
            text=str(text)
            if ':' in text or ',' in text:
                 if '!' in text:
                    text=text[text.find('!')+1:]
                 if ':' in text:   
                     _A_,_B_=text.split(':')
                 elif ',' in text:   
                     _A_,_B_=text.split(',')
                 find_alphaber1= re.findall('[a-z]{1,3}',_A_,re.IGNORECASE)
                 find_alphaber2= re.findall('[a-z]{1,3}',_B_,re.IGNORECASE)
                
                 if find_alphaber1 and find_alphaber2:
                     return alphabet_num[find_alphaber1[0].lower()],alphabet_num[find_alphaber2[0].lower()]
                 elif not find_alphaber1 and find_alphaber2:
                     return 0,alphabet_num[find_alphaber2[0].lower()]
                 elif find_alphaber1 and not find_alphaber2:
                     return alphabet_num[find_alphaber1[0].lower()],0
                 else:
                     
                     return _A_,_B_
                 
            find_alphaber= re.findall('[a-z]{1,3}',text,re.IGNORECASE)
            if find_alphaber:
                return alphabet_num[find_alphaber[0].lower()],None
            return text,None
    def delete_col(self,range_,*sheetID,**kwargs):
               sheetId=  self.__arg_return_sheetid(*sheetID,**kwargs)
               if sheetId==None:return "Not found this sheet name"
               if re.findall('[a-z]',str(range_),re.IGNORECASE):
                   start,end=self._find_columb_alphabet_to_number(range_)
                   if end==None:
                      end= start 
               else:
                   start,end=range_
               if start<1:start=1
               body=   {
                  "requests": [
                    {
                      "deleteDimension": {
                        "range":{
                             "sheetId": sheetId,
                              "dimension": "COLUMNS",
                              "startIndex": int(start)-1,
                              "endIndex": int(end)
                            }}}]}
               request = self.service_sheet.spreadsheets().batchUpdate(spreadsheetId=self._id ,body=body).execute()           
    def append_col(self,length,*sheetID,**kwargs):
               sheetId=  self.__arg_return_sheetid(*sheetID,**kwargs)
               if sheetId==None:return "Not found this sheet name"
               
               body=  {"requests":[{
                      "appendDimension": {
                        "sheetId": sheetId,
                        "dimension": "COLUMNS",
                        "length": length}
                      }
                    ]}
               request = self.service_sheet.spreadsheets().batchUpdate(spreadsheetId=self._id ,body=body).execute()
    def __arg_return_sheetname(self,*sheetID,**kwargs):
        sheetId=0
        sheet_name=''
        if sheetID:
              sheetId=sheetID[0]
        if kwargs:
            if 'sheetname' in kwargs:
                sheet_name=kwargs['sheetname']
                for i in self.sub:
            
                    sheet_id_,name_=self.sub[i]
                    if sheet_name == name_:
                        
                        return name_
                return None
        if not sheet_name:
          for i in self.sub:
            
            sheet_id_,name_=self.sub[i]
            if sheetId == sheet_id_:
                
                return name_
                
        return None
    def __arg_return_sheetid(self,*sheetID,**kwargs):
        sheetId=0
        sheet_name=''
        if sheetID:
              sheetId=sheetID[0]
        if kwargs:
            for j in kwargs:
                
                findid=re.findall('sheetid|id',str(j),re.IGNORECASE)
                if findid:
                    return int(kwargs[j])
            if 'sheetname' in kwargs:
                sheet_name=kwargs['sheetname']
         
                for i in self.sub:
            
                    sheet_id_,name_=self.sub[i]
                    if sheet_name == name_:
                        
                        return sheet_id_
                return None
                
        return sheetId
    
    def sort_col(self,name,*Updown,**col):
            
            
            sheetId=0
            updown=1
            if col:
                if 'col' in col:
                    chose_col=int(col['col'])-1
                    if chose_col<0:
                        chose_col=0
            chose_col_=str(self.get_col(name))
            
            if  chose_col_ and re.findall('\d+',chose_col_,re.IGNORECASE):
                chose_col=int(chose_col_)-1
            else:
                if chose_col:pass
                else:return
            if Updown:
                if Updown[0]==1:updown=1
                else:updown=0
            if updown:
                updown='ASCENDING'
            else:
                updown='DESCENDING'
            
            body=        {
              "requests": [
                {
                  "sortRange": {
                    "range": {
                      "sheetId": sheetId,
                      "startRowIndex": 1,
                      "endRowIndex": 50,
                      "startColumnIndex": 0,
                      "endColumnIndex": 50
                    },
                    "sortSpecs": [
                      {
                        "dimensionIndex": int(chose_col),
                        "sortOrder": updown
                      }]  }
                    }
                  ]
                }
            request = self.service_sheet.spreadsheets().batchUpdate(spreadsheetId=self._id ,body=body).execute()
    
    def get_sheet_size(self,*sheetID,**kwargs): # check specific worksheet Max column and Max row
        sheet_name=self.__arg_return_sheetname(*sheetID,**kwargs)
        response=self.service_sheet.spreadsheets().values().get(spreadsheetId=self._id, range=sheet_name).execute()['range']
        
        index1=response.find('!')
        range1,range2=response[index1+1:].split(':')
       
        max_row=re.findall('\d+',str(range2),re.IGNORECASE)[0]
        max_column=alphabet_num[re.findall('[a-z]{1,3}',str(range2),re.IGNORECASE)[0].lower()]
        return int(max_column),int(max_row)
    def find_string(self,find_sting,*sheetID,**kwargs):
        find_list=[]
        sheet_name=self.__arg_return_sheetname(*sheetID,**kwargs)
        if not sheet_name:return "Not Found"
        response=self.service_sheet.spreadsheets().values().batchGet(spreadsheetId=self._id, ranges=sheet_name ).execute()['valueRanges'][0]['values']
        
        for i in range(0,len(response)):
            for j in range (0,len(response[i])):
               
               if find_sting == str(response[i][j].translate(non_bmp_map)):
                  find_list.append((j+1,i+1))
                  
           
        return find_list
    def get_col_index(self,colname,*sheetID,**kwargs):
        sheet_name=self.__arg_return_sheetname(*sheetID,**kwargs)
        if not sheet_name:return "Not Found "
        response=self.service_sheet.spreadsheets().values().batchGet(spreadsheetId=self._id, ranges=sheet_name ).execute()
        for i in response['valueRanges']:
            for j in range(0,len(i['values'][0])):
               
                    if colname == str(i['values'][0][j].translate(non_bmp_map)):
                                 return j+1
        return "Not Found"
                        
    
    def get_row_index(self,rowname,*sheetID,**kwargs): # return first find row index        
        sheet_name=self.__arg_return_sheetname(*sheetID,**kwargs)
        if not sheet_name:return "Not Found"
        response=self.service_sheet.spreadsheets().values().batchGet(spreadsheetId=self._id, ranges=sheet_name ).execute()['valueRanges'][0]['values']
        
        for i in range(0,len(response)):
            for j in range (0,len(response[i])):
               if rowname == str(response[i][j].translate(non_bmp_map)):
                   return i+1
           
        return "Not Found"
    def __formula__get_col_index(self,colname,*sheetID):
        sheetId=0
        if sheetID:
              sheetId=sheetID[0]
        name_chose=''
        for j in self.sub:
            sheet_id_,name_=self.sub[j]
            if sheetId == sheet_id_:
                name_chose=name_+'!'
                break
        formula='match("'+colname+'",'+str(name_chose)+'a1:in1,0)'
     
       
        self.formula(self.formula_use_data,formula)
        
        
        if re.findall('\d+',self.read(self.formula_use_data),re.IGNORECASE):
          colnum=int(re.findall('\d+',self.read(self.formula_use_data),re.IGNORECASE)[0])
          return colnum
        else:
          print('◭ Error: '+str(colname)+" Was Not Found")
          return None
      
        return int(colnum)
    def __formula__get_col(self,findname,*sheetID):
        sheetId=0
        if sheetID:sheetId=sheetID[0]
        name_chose=''
        for j in self.sub:
            sheet_id_,name_=self.sub[j]
            if sheetId == sheet_id_:
                name_chose=name_+'!'
                break
            
            

        formula='=ArrayFormula(IFERROR(ADDRESS(SMALL(IF(IFERROR(FIND("'+findname+'",'+str(name_chose)+'A1:Z1000),0)>0,column('+str(name_chose)+'A1:Z1000),""),ROW('+str(name_chose)+'1:1)),1,4),""))'
        self.formula(self.formula_use_data,formula)
        if re.findall('\d+',self.read(self.formula_use_data),re.IGNORECASE):
          col=int(re.findall('\d+',self.read(self.formula_use_data),re.IGNORECASE)[0])
          return col
        else:
          print('◭ Error: '+str(findname)+" Was Not Found")
          return None
    def __formula__get_row(self,findname,*sheetID):
        sheetId=0
        if sheetID:sheetId=sheetID[0]
        name_chose=''
        for j in self.sub:
            sheet_id_,name_=self.sub[j]
            if sheetId == sheet_id_:
                name_chose=name_+'!'
                break
            
            

        formula='=ArrayFormula(IFERROR(ADDRESS(SMALL(IF(IFERROR(FIND("'+findname+'",'+str(name_chose)+'A1:Z1000),0)>0,ROW('+str(name_chose)+'A1:Z1000),""),ROW('+str(name_chose)+'1:1)),1,4),""))'
        self.formula(self.formula_use_data,formula)
        if re.findall('\d+',self.read(self.formula_use_data),re.IGNORECASE):
          row=int(re.findall('\d+',self.read(self.formula_use_data),re.IGNORECASE)[0])
          return row
        else:
          print('◭ Error: '+str(findname)+" Was Not Found")
          return None
    
    
    def formula(self,cell,formula,*sheetID):
          sheetId=0
         
          if '!' in str(cell):
            titlename,_____=cell.split('!')
            if 'for_formula_use_' not in str(self.sub):
                self.add_sheet("for_formula_use_")
           
           
            for j in self.sub:
              sheetidd,name_sheet=self.sub[j]
              if str(name_sheet)==str(titlename):
                  sheetId=sheetidd
                  break
          
          
          if not re.findall('=',str(formula),re.IGNORECASE):
              formula='='+formula
          if re.findall("'",str(formula),re.IGNORECASE):
              formula=formula.replace("'",'"')
          if sheetID:
              sheetId=sheetID[0]
           
          
          if ':' in cell: #A1:B1
              if '!'in str(cell):__,cell=cell.split('!')
              input1,input2=cell.split(':')
              find_alp1= re.findall("[a-z]{1,2}",input1,re.IGNORECASE)
              find_alp2= re.findall("[a-z]{1,2}",input2,re.IGNORECASE)
              if find_alp1 and find_alp2:
                  col=alphabet_num[find_alp1[0].lower()]
                  colend=alphabet_num[find_alp2[0].lower()]
                  row=re.findall("\d+",input1,re.IGNORECASE)[0]
                  rowend=re.findall("\d+",input2,re.IGNORECASE)[0]
              else:
                  if '!'in str(cell):__,cell=cell.split('!')
                      
                  col=input1
                  row=input2
                  colend=int(col)+1
                  rowend=int(row)+1
                  
          else:# A5
              if '!'in str(cell):__,cell=cell.split('!')
              find_alp= re.findall("[a-z]{1,3}",cell,re.IGNORECASE)
              if find_alp:
                  col=alphabet_num[find_alp[0].lower()]
                  colend=int(col)+1
                  row=re.findall("\d+",cell,re.IGNORECASE)[0]
                  
                  rowend=int(row)+1
         # formula="=MATCH(\""+findvalue+"\",'A1:IN1',0)"
          
          body=      {
                  "requests": [
                    {
                      "repeatCell": {
                        "range": {
                          "sheetId": sheetId,
                         
                          "startColumnIndex": int(col)-1,
                          "endColumnIndex": int(colend)-1,
                          "startRowIndex": int(row)-1,
                          "endRowIndex":int(rowend)-1
                        },
                        "cell": {
                          "userEnteredValue": {
                              "formulaValue": formula
                          }
                        },
                        "fields": "userEnteredValue"
                      }
                    }
                  ]
                }
          request = self.service_sheet.spreadsheets().batchUpdate(spreadsheetId=self._id ,body=body).execute()
          
    def getsub_id(self,find):
                      
                      for j in self.sub:
                         sheet_id_,name_= self.sub[j]
                         if find==name_:
                             return sheet_id_
                      return ''
    def FNR(self,find_str,replace_str,*sheetID,**kwargs):
        #  if set FNR_allsheet , sheetId must be None
          sheetId=0
          sheetId=self.__arg_return_sheetid(*sheetID,**kwargs)
          FNR_allsheet=0
          
          if kwargs:
             
              if 'id' in kwargs:
                  sheetId=kwargs['id']
                  if not sheetId.isdecimal():
                      print("Sheet ID must be positive number")

                      return
                  else:
                      if int(sheetId)<0:
                          print("Sheet ID must be positive number")
              
                
              if 'allsheet'in kwargs:
                  FNR_allsheet=1
              if 'name' in kwargs:
                  sheet_name=kwargs['name']
                  sheet_id=self.getsub_id(sheet_name)
                  if sheet_id:
                      sheetId=sheet_id
                  
                         
          if FNR_allsheet:
              FNR_allsheet=True
          else:
              FNR_allsheet=False
          
          if FNR_allsheet:    
                  body=      {"requests":[{'findReplace':{
                  "find": find_str,
                  "replacement": replace_str,
                  "matchCase": True ,
                  "matchEntireCell": True ,
                  "searchByRegex": True ,
                  "includeFormulas": True ,
                  "allSheets": True
                  }}]

           
              }
          else:
                  body=      {"requests":[{'findReplace':{
                  "find": find_str,
                  "replacement": replace_str,
                  "matchCase": True ,
                  "matchEntireCell": True ,
                  "searchByRegex": True ,
                  "includeFormulas": True ,
                  "sheetId": sheetId
                  }}]

               
                  }
                  
          request = self.service_sheet.spreadsheets().batchUpdate(spreadsheetId=self._id ,body=body).execute()
    def __range__return(self,text):
        arr=[]
        if isinstance(text,tuple):
            arr.append(text)
            return arr
        text=str(text)
        
        
        if ':' in text:
                _a,_b=text.split(':')
                if re.findall('[a-z]',str(_a),re.IGNORECASE) and re.findall('[a-z]',str(_b),re.IGNORECASE) :
                    colstart,colend=self.___get__index_start_end(text)
                    rowstart=re.findall('\d+',str(_a),re.IGNORECASE)[0]
                    rowend=re.findall('\d+',str(_b),re.IGNORECASE)[0]
                    for i in range(int(colstart),int(colend)+1):
                        for j in range(int(rowstart),int(rowend)+1):
                            arr.append((i,j))
                    
                    
        else:
            find_= re.findall('[a-z]{1,3}',str(text),re.IGNORECASE)
            if find_:
                colstart=alphabet_num[find_[0].lower()]
                rowend=re.findall('\d+',str(text),re.IGNORECASE)[0]
                arr.append((int(colstart),int(rowend)))
        return arr
    def reset_color(self,set_range,**kwargs):
        
        self.setcolor(set_range,(255,255,255),'',**kwargs)
    def ___color_request_body__(self,R,G,B,alpha,col_number,row_number,sheetId):
        
                       body= {
                            "updateCells":{
                            "rows":[{
                                    "values":[{"userEnteredFormat":{
                                               "backgroundColor":
                                               {
                                               "red":   R/255,
                                               "green": G/255,
                                               "blue":  B/255,
                                               "alpha": alpha

                                               }}
                                    }]}],
                            "fields":'userEnteredFormat.backgroundColor',
                            "start":{  # range <==> start(A1) only set 1 cells(sheetId,rowIndex,columnIndex,)
                                    "sheetId": sheetId,
                                    "rowIndex": int(row_number)-1,
                                    "columnIndex": int(col_number)-1
                                    }
                            }
                      }
                       return body
   
    
    def set_color_PIL(self,ndarray,start_x,start_y,reqs_range,*args,**kwargs):# if want to use HEXcolor RGB set 0
            
            body={}
            body["requests"]=[]
            sheetId=0 # ID default  workbook1
            alpha=1   # Transparent default 1
            #ndarray=cv2.resize(ndarray, (400, 400), interpolation=cv2.INTER_CUBIC)
            #height,width,_=ndarray.shape
            width,height=ndarray.size
            runpixel=0
            
            sheetId=  self.__arg_return_sheetid(0,**kwargs)
            
            for x in range(start_x,width+1):
                for y in range(start_y,height+1):
                    
                    
                    R,G,B=ndarray.getpixel((x-1,y-1))
                  
                    
                    
                    #R,G,B=ndarray[x-1,y-1]
                    body["requests"].append(self.___color_request_body__(R,G,B,alpha,x,y,sheetId))
                   
                    runpixel+=1
                    if runpixel>reqs_range:break
                   
                if y >=height and x<=width+1:
                    start_y=1
                if runpixel>reqs_range:break
                
                    
            
            max_col,max_row=self.get_sheet_size(sheetId)
            
      
            addcol=0
            addrow=0
            if int(width)>int(max_col):
                addcol=int(width)-int(max_col)
                self.append_col(addcol,sheetId)
            else:
                delcol=abs(int(width)-int(max_col))
                self.delete_col((0,delcol),sheetId)
            if int(height)>int(max_row):
                addrow=int(height)-int(max_row)
                self.append_row(addrow,sheetId)
            else:
                delrow=abs(int(height)-int(max_row))
                self.delete_row((0,delrow),sheetId)
            self.adjust_col_row('all',1,'all',1,sheetId)
            
            if kwargs:
                find_alpha=re.findall('alpha|transparent',str(kwargs),re.IGNORECASE)
                if find_alpha and len(find_alpha)<2:
                    
                    alpha=int(kwargs[find_alpha[0]])
                find_id=re.findall('sheetid|id',str(kwargs),re.IGNORECASE)
                if find_id and len(find_id)<2:
                   
                    sheetId=int(kwargs[find_id[0]])
            
            
            request = self.service_sheet.spreadsheets().batchUpdate(spreadsheetId=self._id ,body=body).execute()
            
            time.sleep(0.5)
            print("Next Painting...")
            if x>=width and y>=height:return ndarray
            elif y>=height and x<=width:
                y=1
                x+=1
            else:
                y+=1
            self.set_color_PIL(ndarray,x,y,reqs_range)
    def set_color_ndarray(self,ndarray,start_x,start_y,*args,**kwargs):# if want to use HEXcolor RGB set 0
            
            body={}
            body["requests"]=[]
            sheetId=0 # ID default  workbook1
            alpha=1   # Transparent default 1
            #ndarray=cv2.resize(ndarray, (400, 400), interpolation=cv2.INTER_CUBIC)
            height,width,_=ndarray.shape
            runpixel=0
            
            sheetId=  self.__arg_return_sheetid(0,**kwargs)
            
            for x in range(start_x,width+1):
                for y in range(start_y,height+1):
                    runpixel+=1
                    if runpixel>2500:break
            
                    R,G,B=ndarray[x-1,y-1]
                    body["requests"].append(self.___color_request_body__(R,G,B,alpha,x,y,sheetId))
                if runpixel>2500:break
           
            
            
            
            max_col,max_row=self.get_sheet_size(sheetId)
      
            addcol=0
            addrow=0
            if int(width)>int(max_col):  addcol=int(width)-int(max_col)
            if int(height)>int(max_row):  addrow=int(height)-int(max_row)
            
            if addcol: self.append_col(addcol,sheetId)
            if addrow: self.append_row(addrow,sheetId)
            self.adjust_col_row('all',10,'all',10,sheetId)
            
            if kwargs:
                find_alpha=re.findall('alpha|transparent',str(kwargs),re.IGNORECASE)
                if find_alpha and len(find_alpha)<2:
                    
                    alpha=int(kwargs[find_alpha[0]])
                find_id=re.findall('sheetid|id',str(kwargs),re.IGNORECASE)
                if find_id and len(find_id)<2:
                   
                    sheetId=int(kwargs[find_id[0]])
            
            
            request = self.service_sheet.spreadsheets().batchUpdate(spreadsheetId=self._id ,body=body).execute()
            if x>=width and y>=height:return ndarray
            self.set_color_ndarray(ndarray,x,y)
    def setcolor(self,set_range,RGB,*args,**kwargs):# if want to use HEXcolor RGB set 0
            
            coordinate_arr=self.__range__return(set_range)
           
            if not coordinate_arr:return
            sheetId=0 # ID default  workbook1
            alpha=1   # Transparent default 1
            sheetId=  self.__arg_return_sheetid(0,**kwargs)
            
            
            if kwargs:
                find_alpha=re.findall('alpha|transparent',str(kwargs),re.IGNORECASE)
                if find_alpha and len(find_alpha)<2:
                    
                    alpha=int(kwargs[find_alpha[0]])
                find_id=re.findall('sheetid|id',str(kwargs),re.IGNORECASE)
                if find_id and len(find_id)<2:
                   
                    sheetId=int(kwargs[find_id[0]])
                    
            if 'tuple'  in str(type(RGB)):
                R,G,B=RGB
            if args :
                if '#' in str(args[0]) and 'tuple' not in str(type(RGB)): 
                    R,G,B=HEX_to_RGB(args[0])
            body={}
            body["requests"]=[]
            for each in coordinate_arr:
                coll,roww=each
                body["requests"].append(self.___color_request_body__(R,G,B,alpha,coll,roww,sheetId))
                
            
           
            request = self.service_sheet.spreadsheets().batchUpdate(spreadsheetId=self._id ,body=body).execute()
                  
    @property
    def id(self):
        return self._id
    @id.setter
    def id(self,newid):
        self._id=newid
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
            
            
        
            
            request = self.service_sheet.spreadsheets().get(spreadsheetId=self._id, includeGridData=False)
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
    def add_sheet(self,newpagename,*args):   #create　new page in exist spreadsheet
        
        
        request_=[{"addSheet":{"properties":{"title": newpagename}}}]
        if  args:
            for new_page in args:
                request_.append({"addSheet":{"properties":{"title": new_page}}})
        body={"requests":request_}
        request = self.service_sheet.spreadsheets().batchUpdate(spreadsheetId=self._id,body=body)
        request.execute()
        self.getsheet()
    def delete_sheet(self,delete_pagename,*args):   #create　new page in exist spreadsheet
                
                request_=[]
                for i in self.sub:
                    sheet_id_,name_page=self.sub[i]
                    if delete_pagename in name_page:
                        request_.append({"deleteSheet":{"sheetId":sheet_id_}})
                
                
                        
                body={"requests":request_}
                request = self.service_sheet.spreadsheets().batchUpdate(spreadsheetId=self._id,body=body)
                request.execute()
                self.getsheet()
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
        request = self.service_sheet.spreadsheets().batchUpdate(spreadsheetId=self._id ,body=body)
        request.execute()
    def rename_workbook(self,sheetId,newname):
        body={"requests":[{
            "updateSheetProperties":{
                "properties":{
                    "sheetId":sheetId,"title": newname},"fields": 'title'}}]}
        request = self.service_sheet.spreadsheets().batchUpdate(spreadsheetId=self._id ,body=body)
        request.execute()
    @classmethod    
    def create(cls,filename,*args,**kwargs):  #create　new spreadsheet
    
        sheet_page=[]
        if kwargs:
                    if 'folder' in kwargs:
                            foldername=kwargs['folder']
                            sheetid=Drive().create_newsheet(filename,folder=foldername)
                            
                            cls.id=sheetid
                            
                            cls.getsheet(cls)
                            return cls.id
        if  args:
            for new_page in args:
                sheet_page.append({"properties":{"title": new_page}})
        body= {"properties": {"title":filename},"sheets":sheet_page }
        request = self.service_sheet.spreadsheets().create(body=body)
        response=request.execute()
        
        
       # self._id=response['spreadsheetId']
       # cls.id=self._id
        cls._id=response['spreadsheetId']
        cls.getsheet(cls)
        
        return cls._id
    def copy(self,sheetId):             # copy a new workbook in exist spreadsheet
        body={ 'destination_spreadsheet_id': self._id}
        
        request = self.service_sheet.spreadsheets().sheets().copyTo(spreadsheetId=self._id,sheetId=sheetId,body=body)
        request.execute()
    def read(self,workRange):
        response=self.service_sheet.spreadsheets().values().batchGet(spreadsheetId=self._id, ranges=workRange ).execute()
        try:
            if len(response.get('valueRanges')[0]['values'][0])==1 and len(response.get('valueRanges')[0]['values'])==1:
                
                    return response.get('valueRanges')[0]['values'][0][0]
            else:
                    return response.get('valueRanges')[0]['values']
        except:
            return ''
    def delete(self,workRange,*args):
        self.write(workRange,'')
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
    
   
    
        request = self.service_sheet.spreadsheets().values().update(spreadsheetId=self._id, range=workRange,valueInputOption='RAW', body=writeData )
        response = request.execute()
class Gmail:
    tokenpickle=''
    json=''
    def __init__(self,*args,**kwargs):
        global service_mail
        if args:
            find_cred=re.findall('token|pickle|json',str(args[0]),re.IGNORECASE)
            if find_cred:
                if re.findall('token|pickle',str(find_cred[0]),re.IGNORECASE):
                          self.tokenpickle=args[0]
                else:
                          self.json=args[0]
                service_mail=self.main()
        elif kwargs:
            if 'token' in kwargs :
                self.tokenpickle=kwargs['token']
            if 'pickle' in kwargs :
                self.tokenpickle=kwargs['pickle']
            if 'json' in kwargs :
                self.json=kwargs['json']
            service_mail=self.main()
    def list_all(self,*args,**kwargs):
        print_=1
        if 'print_' in kwargs:
            print_=kwargs['print_']
        if 'dict_'in kwargs:
            dict_all=kwargs['dict_']
        if not args:
            dict_all={}
            pageToken=None
            count=0
            
        else:
                pageToken=args[0]
                count=  args[1]
       
        result=service_mail.users().messages().list(userId='me',pageToken=pageToken).execute()
        if 'nextPageToken' not in result or 'messages' not in result:
            return dict_all
            
        result_msg=result['messages']
        for i_d in result_msg:
            count+=1
            if not  print_:    print(str(count),'.'+i_d['id'])
            message = service_mail.users().messages().get(userId='me', id=i_d['id'],
                                             format='full').execute()
            if  print_:
                for i in message['payload']['headers']:
                    if 'subject' in str(i['name']).lower():
                        subject=i.get('value')
                    if 'from' in str(i['name']).lower():
                        from_=i.get('value')
                    if 'to' in str(i['name']).lower():
                        to=i.get('value').split(',')[0]
                ts=message['internalDate'][:10]
                sendtime=datetime.datetime.fromtimestamp(int(ts)).strftime('%Y-%m-%d %H:%M:%S')
                print(str(count),'.'+i_d['id'])
                print("Sub  :",   subject.translate(non_bmp_map))
                print("ID   :",i_d['id'])
                print("From :",from_)
                print("To   :",to)
                print("Recv Time:",sendtime)
                print("=====================")
        return self.list_all(result['nextPageToken'],count,dict_=dict_all,print_=print_)
    def get_attach_data(self,*args):
         if arg:
             msg_id,attch_id=arg[0],arg[1]
             
         data =  service_mail.users().messages().attachments().get(userId='me',messageId=msg_id, id=attch_id).execute()
         return data
    def read(self,*args):
            data=''
            if args:
                id=args[0]
                message = service_mail.users().messages().get(userId='me', id=id,
                                             format='full').execute()
            for i in message['payload']['headers']:
                if 'subject' in str(i['name']).lower():
                    subject=i.get('value')
                if 'from' in str(i['name']).lower():
                    from_=i.get('value')
                if 'to' in str(i['name']).lower():
                    to=i.get('value').split(',')[0]
            ts=message['internalDate'][:10]
            sendtime=datetime.datetime.fromtimestamp(int(ts)).strftime('%Y-%m-%d %H:%M:%S')
            if 'data' not in message['payload']['body']:
               list_=message['payload']['parts']
               if len (list_)==1:
                   data=message['payload']['parts'][0]['body']['data']
               else:
                   for _i_ in range(0,len(list_)):
                       data_=message['payload']['parts'][_i_]['body']['data']
                       data+=data_
            else:    
               data=message['payload']['body']['data']
            data = data.replace("-","+").replace("_","/")
            decoded_data = base64.b64decode(data)
            soup = BeautifulSoup(decoded_data , "lxml")
            print("Sub  :",   subject.translate(non_bmp_map))
            print('Snap :',message['snippet'])
            print("ID   :",   id)
            print("From :",from_)
            print("To   :",to)
            print("Recv Time:",sendtime)
            print("Content: \n",soup.text.replace('\u3000',''))
            print("=====================")
    def main(self):
        """Shows basic usage of the Drive v3 API.
        Prints the names and ids of the first 10 files the user has access to.
        """
        creds = None
        service_account=0
        # The file token.pickle stores the user's access and refresh tokens, and is
        # created automatically when the authorization flow completes for the first
        # time.
        
        if os.path.exists(self.tokenpickle):
            
            with open(self.tokenpickle, 'rb') as token:
                creds = pickle.load(token)
         
        # If there are no (valid) credentials available, let the user log in.
        try:
         if not creds or not creds.valid:
            
            if creds and creds.expired and creds.refresh_token:
                creds.refresh(Request())
            else:
               
                if os.path.exists(self.json):
                 with open(self.json,'r') as jsooon:
                    if 'service_account' in str(jsooon.read()):
                        service_account=1
                if service_account:
                    creds = ServiceAccountCredentials.from_json_keyfile_name(
                        self.json, scopes=Mail_SCOPES)
                    
                    
                else:
                    flow = InstalledAppFlow.from_client_secrets_file(
                        self.json, Mail_SCOPES)
                    creds = flow.run_local_server(port=8012)
                   
                    # Save the credentials for the next run
                    self.tokenpickle='token.pickle'
                    with open('token.pickle', 'wb') as token:
                
                        pickle.dump(creds, token)
        except Exception as Errorlogin:
            print(Errorlogin)
            print("▲ If you don't set the json or pickle path, plz set the path before use ")
            print("Use function: \n 1. chose_pickle (pickle_path)\n 2. chose_json   (json_path)")
            return
        
        service_mail = build('gmail', 'v1', credentials=creds)  
        return service_mail
        
