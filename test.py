import gdapi
import time

start=time.time()

id_='1BP0c4be04rqva7pSfrqGrmBhJwVN7lScxz5mBj3d6dk'
sheet=gdapi.Writer(id_)
count=1
for i in range(0,100):
        
        range_='B'+str(count)
      
        sheet.write(range_,'Hello world')
        count+=1
        time.sleep(1)
print(time.time()-start)
ads+=1

# if you want to create a new sheet

file_id=gdapi.create_newsheet("new_sheet_name") # get a name of new_sheet_name Sheet

# change permission

gdapi.change_permissions(file_id) # 

# Write Sheet
sheet=gdapi.Writer(file_id)

sheet.write('B5','In new_sheet_name, Hi this is B5')
