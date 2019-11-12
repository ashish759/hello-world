#Need to install requests package for python
#sudo easy_install requests
import requests
import xlrd 
 
# Set the request parameters
url = 'https://dev55662.service-now.com/incident_list.do?EXCEL&sysparm_default_export_fields=sys_id,number,short_description,assignment_group,priority,description,assigned_to&sysparm_display_value=true&sysparm_exclude_reference_link=true&sysparm_query=active=true^assignment_groupLIKE(CG)'
user = 'admin'
pwd = 'Ashi@1234'
 
# Set proper headers
headers =  {"Content-Type":"application/json","Accept":"application/xml"}
 
# Do the HTTP request
print (url) 
r = requests.get(url, auth=(user, pwd))  
page = r.content
print ("downloaded")
f=open(r'C:/Users/ashishkp/Downloads/incident.xls', 'wb')
f.write(page)
f.close()
path = f.name
path = (path.replace("\\","/"))
print (path)
#read incident excel(dwnloaded excel)
loc = ("C:/Users/ashishkp/Downloads/incident.xls") 
  
wb1 = xlrd.open_workbook(loc) 
sheet1 = wb1.sheet_by_index(0)

sheet1.cell_value(0, 0)

loc2 = ("C:/Users/ashishkp/Desktop/NewPrograms/incident.xlsx") 
      
wb2 = xlrd.open_workbook(loc2) 
sheet2 = wb2.sheet_by_index(0)

sheet2.cell_value(0, 0 )


#store the number value


  
for j in range(sheet1.nrows):
     if cell.value == "Additional comments":
         number_val=cell.column
         for i in range(1,max_row):
            num_counter = str(i)+str(number_val)
            num_val=ws[num_counter].value
            print(num_val)
            um_cntr = str(i)+str(umber_val)
            um_val=ws[um_cntr].value
            print(um_val)
    
    
     
     
     
    
                   
   
    
    














 

       

#store the assignment grp value



              
                 
        
       
    
       
    
    
    
 
#read mapping sheet

  

     
    #check assignment grp == mapping sheet_assign grp:


                
    #pull team lead value
   
                
##    print(sheet2.cell_value(i, j+1))




