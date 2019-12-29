# -*- coding: utf-8 -*-
"""
Created on Sat Dec 14 22:18:05 2019

@author: hp
"""


import pandas as pd 

from tkinter import *





import sys, os 
def resource_path(relative_path):
    if hasattr(sys, '_MEIPASS'):
        return os.path.join(sys._MEIPASS, relative_path)
    return os.path.join(os.path.abspath("."), relative_path)








data = pd.read_csv(resource_path("teacher_courses.csv")) 


OPTIONS = data.iloc[:,0].values
OPTIONS = list(OPTIONS)

OPTIONS = [x for x in OPTIONS if str(x) != 'nan']

OPTIONS2 = data.iloc[:,1].values





master = Tk()

master.title("FEEDBACK: AID")



w3 = Canvas(master, width=500, height=100) 
w3.pack() 


variable = StringVar(master)
variable.set(OPTIONS[0]) # default value


variable2 = StringVar(master)
variable2.set(OPTIONS2[0]) # default value


w = OptionMenu(master, variable, *OPTIONS)
w.pack()


w2 = OptionMenu(master, variable2, *OPTIONS2)
w2.pack()



w3 = Canvas(master, width=500, height=50) 
w3.pack() 







def ok():
    #print ("value is:" + variable.get())
    #print ("value is:" + variable2.get())
    
    
    button.destroy()
    w4.destroy()
    text2 = Text(master,wrap=WORD, height=3, width=30)
    text2.insert(INSERT, "Please Wait..........")
    text2.pack()
    w3 = Canvas(master, width=500, height=67) 
    w3.pack() 
    
    
    
    import time
    progress['value']= 5
    master.update_idletasks()
    time.sleep(1)
    progress['value']=10
    master.update_idletasks()
    time.sleep(1)
    progress['value']=15
    master.update_idletasks()
    time.sleep(2)
    progress['value']=25
    master.update_idletasks()
    time.sleep(1)
    
    
    d= pd.read_csv(resource_path("feedback.csv"))

    name = variable.get()#"Dr. SK Singh"#input("Input faculty name: ")
    course = variable2.get()#"Microprocessors"#input("Input course name: ")
    session = "July-Dec"##input("Input course session: ")
    
    
    
    x = d.iloc[  :  , 1: ]
    y = d.iloc[ :  , -1].values
    question = d.columns
    
    
    
    ## col 1 to 5
    a = [[0,0,0,0,0],[0,0,0,0,0],[0,0,0,0,0],[0,0,0,0,0],[0,0,0,0,0],[0,0,0,0,0],[0,0,0,0,0],[0,0,0,0,0],[0,0,0,0,0],[0,0,0,0,0],[0,0,0,0,0],[0,0,0,0,0]]  #[no. of questions][1-5 marks]
    
    k = 0
    for j in range(7,19):
        
        for i in range(len(x)):
            if d.iloc[i, 3] == course and d.iloc[i,4] == name:
                
                a[k][d.iloc[i , j]-1] = a[k][d.iloc[i , j]-1] + 1
                
            else:
                continue
            
        k = k+1
        
    #print(a)
    
    
    sum2 = [0,0,0,0,0] #col 1 to 5
    b = list(map(list, zip(*a)))
    
    for i in range(len(b)):
                
            sum2[i] = sum(b[i])
        
    #print(sum2)
        
    
    
    ## col total marks
    
    sum1 = [0,0,0,0,0,0,0,0,0,0,0,0]  ## no of questions
    
    
    for i in range(len(a)):
        for j in range(len(a[i])):
            
            sum1[i] = sum1[i] + a[i][j]*(j+1) 
            
    
    #print(sum1)
    
    #print(sum(sum1))
    
    
    
    ## col % marks
    
    total_students = sum((a[0]))
    
    per_marks = [0,0,0,0,0,0,0,0,0,0,0,0] ## no of questions
    
    if total_students != 0:
    
        for i in range(len(sum1)):
            
            per_marks[i] = (sum1[i]/(total_students*5))*100
            
    else:
        text2.delete('1.0', END)
        text2.insert(INSERT, "Data is not available for either " + name + " or " + course)
        
        return 
        
    #print(per_marks)
    
    
    
    
    total_marks = sum1
    #col_1to5 = a
    last_row = sum2
    percent_marks = per_marks
    question = question[7:19]
    
    
    
    import xlsxwriter 
    
    file_name = name + "_" + course + ".xlsx"
    
    workbook = xlsxwriter.Workbook(file_name) 
    worksheet = workbook.add_worksheet() 
    
    worksheet.write(0, 0, "Name of teacher")
    worksheet.write(0, 1, name)
    
    worksheet.write(1, 0, "Subject")
    worksheet.write(1, 1, course)
    
    worksheet.write(2, 0, "Session")
    worksheet.write(2, 1, session)
    
    worksheet.write(3, 0, "Total Students")
    worksheet.write(3, 1, total_students)
    
    
    head = ["Sr. No.", "Description", "Very poor", "Poor", "Good", "Very Good", "Excellent", "Total Marks", "%Marks"]
    
    column = 0
    
    for i in head:
        worksheet.write(7, column, i)
        column += 1
    
    
    
    row = 8
    column = 0
      
    
    for i in range(len(question)):
         worksheet.write(row, 0, i+1)
         worksheet.write(row, 1, question[i])
         worksheet.write(row, 7, total_marks[i])
         worksheet.write(row, 8, percent_marks[i])
         
         row += 1
    
    row = 8
    column = 2
    
    for i in range(len(a)):
        for j in range(len(a[i])):
            
            worksheet.write(row, column, a[i][j])
            column += 1
            
        row += 1
        column = 2
    
    worksheet.write(row, 1, "Total")
    column = 2  
    
    for i in last_row:
        worksheet.write(row, column, i)
        column += 1
    
    worksheet.write(row, column, sum(total_marks))   
        
    
    
    
    
    avg_feed_marks = sum(total_marks)/len(question)
    max_marks = total_students*5
    per_avg_marks = 0
    
    if total_students != 0:
        per_avg_marks = (avg_feed_marks/max_marks)*100
    
    avg_feed_scale25 = (per_avg_marks*25)/100
    
    
    
    
    
    
    progress['value']=80
    master.update_idletasks()
    time.sleep(1)
    progress['value']=85
    master.update_idletasks()
    time.sleep(1)
    progress['value']=90
    master.update_idletasks()
    
    
     
    
    
    
    l = [avg_feed_marks,max_marks, per_avg_marks, avg_feed_scale25]
    ll = ["Average Feedback Marks", "Max Marks", "%Average Marks", "Average FeedBack on the Scale of 25"]
    
    column = 0
    row += 2
    for i in range(len(l)):
        worksheet.write(row, column, ll[i])
        worksheet.write(row, column+1, l[i])
        row += 1
    
    
         
    workbook.close() 
    
    
    time.sleep(1)
    progress['value']=100
    
    #button.destroy()
    #w4.destroy()
    w3.destroy()
    text2.destroy()
    text = Text(master,height=1, width=30)
    text.insert(INSERT, "Files Generated..........")
    text.pack()
    
    w3 = Canvas(master, width=500, height=100) 
    w3.pack() 
    
    ## Graph
    
    
    import matplotlib.pyplot as plt 
    
    
    
    plt.bar(question, percent_marks)
    
    x = [per_avg_marks]*len(question)
    
    plt.plot(question,x, color="red", label= "Average") 
    
    plt.title("Average % marks")
    
    plt.xticks(rotation=90)
    
    
    
    plt.savefig(name + "_" + course + ".pdf" ,bbox_inches = "tight")
    
    
    
    
    
    
    
from tkinter.ttk import *
progress=Progressbar(master,orient=HORIZONTAL,length=300,mode='determinate')  
    
progress.pack()


w3 = Canvas(master, width=500, height=10) 
w3.pack() 

button = Button(master, text="SUBMIT", command=ok)
button.pack()


w4 = Canvas(master, width=500, height=100) 
w4.pack() 

mainloop()

























