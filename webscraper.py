from selenium import webdriver
import pandas as pd
import xlsxwriter
import time
def all_zero(date_y,sub):
  for i in sub:
    if date_y[i]!=0:
      return True
  return False

def extract_date(date_x,date_y,sub):
  date = []
  for i in range(113):
      if all_zero(date_y[i],sub):
         date.append(date_x[i])
  return date

def day_wise(date_x,date_y,sub,browser,delay,roll):
  no =""
  if roll<10:
     no = "0"+str(roll)
  else:
     no = str(roll)
  usernameStr = '19MCA'+no
  passwordStr = '19MCA'+no


  time.sleep(delay)
  browser.get(('https://ams.xyz.edu.in/ams/index.php'))
   
  #sending username and password data to the corres field by id
  username = browser.find_element_by_id('user_name')
  username.send_keys(usernameStr)

  password = browser.find_element_by_id('pswd')
  password.send_keys(passwordStr)

  #clicking on the sign in button
  signInButton = browser.find_element_by_id('submit_frm')
  signInButton.click()

  #after 5 seconds loading attendence page
  time.sleep(delay)
  browser.get(('https://ams.xyz.edu.in/ams/index.php?'
             'm=1&page=stu_att_view'))

  #after 5 seconds extracting and displaying the attendence
  time.sleep(delay)
  data = browser.find_elements_by_class_name('table-responsive')
  day_by_day = data[0].find_elements_by_tag_name('td')
  date = ""
  status = 0
  subject = ""
  for entry in day_by_day:
      text = entry.text
      if "2019" in text or "2020" in text:
         date = text
         date_x.append(date)
      if "[P]" in text or "[A]" in text:
         temp = text.split('\n')
         status = 1  
         if temp[1] in sub:
            if  date_y[len(date_x)-1][temp[1]]==0:
                date_y[len(date_x)-1][temp[1]] = status
            else:
                date_y[len(date_x)-1][temp[1]] +=status
      status = 0
  #after 5 seconds loging out of the account 
  time.sleep(delay) 
  browser.get(('https://ams.xyz.edu.in/ams/includes/logout.php'))

def day_wise2(date_y,sub,browser,delay,roll):
  no =""
  date_x=[]
  if roll<10:
     no = "0"+str(roll)
  else:
     no = str(roll)
  usernameStr = '19MCA'+no
  passwordStr = '19MCA'+no


  time.sleep(delay)
  browser.get(('https://ams.xyz.edu.in/ams/index.php'))
   
  #sending username and password data to the corres field by id
  username = browser.find_element_by_id('user_name')
  username.send_keys(usernameStr)

  password = browser.find_element_by_id('pswd')
  password.send_keys(passwordStr)

  #clicking on the sign in button
  signInButton = browser.find_element_by_id('submit_frm')
  signInButton.click()

  #after 5 seconds loading attendence page
  time.sleep(delay)
  browser.get(('https://ams.xyz.edu.in/ams/index.php?'
             'm=1&page=stu_att_view'))

  #after 5 seconds extracting and displaying the attendence
  time.sleep(delay)
  data = browser.find_elements_by_class_name('table-responsive')
  day_by_day = data[0].find_elements_by_tag_name('td')
  date = ""
  status = 0
  subject = ""
  for entry in day_by_day:
      text = entry.text
      if "2019" in text or "2020" in text:
         date = text
         date_x.append(date)
      if "[P]" in text or "[A]" in text:
         temp = text.split('\n')
         if "[P]" in temp:
            status = 1  
         if temp[1] in sub:
            if  date_y[len(date_x)-1][temp[1]]==0:
                date_y[len(date_x)-1][temp[1]] = status
            else:
                date_y[len(date_x)-1][temp[1]] +=status
      status = 0
  #after 5 seconds loging out of the account 
  time.sleep(delay) 
  browser.get(('https://ams.xyz.edu.in/ams/includes/logout.php'))
      
def make_Xl(student):
  s1,s2,s3,s4,s5,s6,s7=[],[],[],[],[],[],[]
  for i in range(1,24):
    if i==6 or i==18:
      continue
    s1.append(student[i]["S1"])
    s2.append(student[i]["S2"])
    s3.append(student[i]["S3"])
    s4.append(student[i]["S4"])
    s5.append(student[i]["S5"])
    s6.append(student[i]["S6"])
    s7.append(student[i]["S7"])
    
  #Create a Pandas dataframe
  df = pd.DataFrame({'S1':s1,'S2':s2,'S3':s3,'S4':s4,'S5':s5,'S6':s6,'S7':s7,})

  # Create a Pandas Excel writer using XlsxWriter as the engine.
  writer = pd.ExcelWriter("FirstYearMCA.xlsx", engine='xlsxwriter')

  # Convert the dataframe to an XlsxWriter Excel object.
  df.to_excel(writer, sheet_name='Sem2_Attendence')

  # Close the Pandas Excel writer and output the Excel file.
  writer.save()

def make_Xl2(date_x,date_z,date_y,sub):
  s1,s2,s3,s4,s5,s6,s7=[],[],[],[],[],[],[]
  date = []
  for i in range(114):
      if date_x[i] in date_z:
         date.append(date_x[i])
         s1.append(date_y[i][sub[0]])
         s2.append(date_y[i][sub[1]])
         s3.append(date_y[i][sub[2]])
         s4.append(date_y[i][sub[3]])
         s5.append(date_y[i][sub[4]])
         s6.append(date_y[i][sub[5]])
         s7.append(date_y[i][sub[6]])
    
  #Create a Pandas dataframe
  df = pd.DataFrame({'date':date,'S1':s1,'S2':s2,'S3':s3,'S4':s4,'S5':s5,'S6':s6,'S7':s7,})

  # Create a Pandas Excel writer using XlsxWriter as the engine.
  writer = pd.ExcelWriter("FirstYearMCA1.xlsx", engine='xlsxwriter')

  # Convert the dataframe to an XlsxWriter Excel object.
  df.to_excel(writer, sheet_name='Sem2_Attendence')

  # Close the Pandas Excel writer and output the Excel file.
  writer.save()
  

def Att_ext(student,date_y,sub,browser,delay):
  for i in range(1,5):
      if i==6 or i==18:
         continue
      if i<10:
         no = "0"+str(i)  
      else:
         no = str(i)
     
      usernameStr = '19MCA'+no
      passwordStr = '19MCA'+no

      time.sleep(delay)
      browser.get(('https://ams.xyz.edu.in/ams/index.php'))
   
      #sending username and password data to the corres field by id
      username = browser.find_element_by_id('user_name')
      username.send_keys(usernameStr)

      password = browser.find_element_by_id('pswd')
      password.send_keys(passwordStr)


      #clicking on the sign in button
      signInButton = browser.find_element_by_id('submit_frm')
      signInButton.click()

      #after 5 seconds loading attendence page
      time.sleep(delay)
      browser.get(('https://ams.xyz.edu.in/ams/index.php?'
             'm=1&page=stu_att_view'))

      #after 5 seconds extracting and displaying the attendence
      time.sleep(delay)
      data = browser.find_elements_by_class_name('table-responsive')
      
      datalist = data[1].find_elements_by_tag_name('td') 
      #going through the attendence
      j = 0
      for item in datalist:
         text = item.text
         if "Over" in text:
           pass
         else:
           j+=1
           temp = text.split(" ")
           student[i]["S"+str(j)]=float(temp[2][1:-2])
      
      #after 5 seconds loging out of the account 
      time.sleep(delay) 
      browser.get(('https://ams.xyz.edu.in/ams/includes/logout.php'))
      
if __name__== "__main__":
   date_x = []
   student = [{} for i in range(24)]
   date_y = [{"[MCA 4442]":0,"[MCA 4444]":0,"[MCA 4446]":0,"[MCA 4452]":0,"[MCA 4450]":0,"[MCA 4648]":0,"[MCA 4454]":0} for i in range(113)]
   sub = ["[MCA 4442]" ,"[MCA 4444]" ,"[MCA 4446]" ,"[MCA 4452]","[MCA 4450]" ,"[MCA 4648]" ,"[MCA 4454]"]
   browser = webdriver.Firefox()
   delay = 3
   Att_ext(student,date_y,sub,browser,delay)
   make_Xl(student)
   '''day_wise(date_x,date_y,sub,browser,delay,1)
   date_z = extract_date(date_x,date_y,sub)
   date_y = [{"[MCA 4442]":0,"[MCA 4444]":0,"[MCA 4446]":0,"[MCA 4452]":0,"[MCA 4450]":0,"[MCA 4648]":0,"[MCA 4454]":0} for i in range(113)]
   day_wise2(date_y,sub,browser,delay,2)
   day_wise2(date_y,sub,browser,delay,4)
   day_wise2(date_y,sub,browser,delay,9)
   day_wise2(date_y,sub,browser,delay,14)
   day_wise2(date_y,sub,browser,delay,17)
   day_wise2(date_y,sub,browser,delay,19)
   make_Xl2(date_x,date_z,date_y,sub)'''
   
   
   
