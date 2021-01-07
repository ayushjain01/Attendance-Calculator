from tkinter import *
from tkinter import filedialog
import xlrd
from datetime import datetime, date
import webbrowser

"""
Credits to Piyush Sir (www.here2share.org) for helping me out with this project.

This project helps schools to calculate attendances easily for the online classes.
Platform like Zoom and Teams allows the teachers to download an Attendance Report, which contains the name of the student, their joining and exiting tiem.
This Attendance Report is a .csv file. 

PRE REQUISITES :
1. Student Database
2. Attendance Report

In case you do not have these files, you can use 'file_gen.py' to generate sample files and run this Project.

Visit :
https://support.zoom.us/hc/en-us/articles/201363213-Getting-started-with-reports  (Zoom) 
or https://support.microsoft.com/en-gb/office/download-attendance-reports-in-teams-ae7cf170-530c-47d3-84c1-3aedac74d310 (Teams)
to learn how to generate reports on Zoom and Teams
"""

def createStudent(book1):
    students = []  # list of all the students
    sheet1 = book1.sheet_by_index(0)
    row, column = 1, 0
    for i in range(sheet1.nrows - 1):
        students.append(sheet1.cell_value(row, column))
        row += 1
    return students

def classdb(studb):
    class6 = []
    class7 =[]
    class8 = []
    class9 = []
    class10 = []
    class11 = []
    class12 = []
    for i in studb:
        if "6" in i:
            class6.append(i)
        elif "7" in i:
            class7.append(i)
        elif "8" in i:
            class8.append(i)
        elif "9" in i:
            class9.append(i)
        elif "10" in i:
            class10.append(i)
        elif "11" in i:
            class11.append(i)
        elif "12" in i:
            class12.append(i)
    return [class6,class7,class8,class9, class10, class11,class12]


def presentlist(book2):
    presentwithdup = []
    sheet1 = book2.sheet_by_index(0)
    row, column = 0, 0
    for i in range(sheet1.nrows):
        presentwithdup.append(sheet1.cell_value(row, column))
        row += 1
    return presentwithdup


def removedup(presentdup):
    uniquelist = []
    for i in presentdup:
        if i not in uniquelist:
            uniquelist.append(i)
    return uniquelist


def absentlist(present, gradelocal):
    absent = []
    for i in gradelocal:
        if i not in present:
            absent.append(i)
    return absent

def time_diff(strtime1, strtime2):
    datetime_object1 = datetime.strptime(strtime1, '%H:%M:%S')
    datetime_object2 = datetime.strptime(strtime2, '%H:%M:%S')
    return int((datetime_object2 - datetime_object1).seconds / 60)

def latejoin(starttime, book2):
    global delay_time
    sheet1 = book2.sheet_by_index(0)
    repeat = 0  # value 0 signifies that name is not repeated yet
    connect = {}
    late = {}
    row, col = 0, 0
    for i in range(sheet1.nrows - 1):
        if repeat == 0:
            connect[sheet1.cell_value(row, col)] = sheet1.cell_value(row, col + 3)
        if sheet1.cell_value(row + 1, col) == sheet1.cell_value(row, col):
            repeat = 1
        else:
            repeat = 0
        row += 1

    if sheet1.cell_value(row - 1, col) != sheet1.cell_value(row, col):
        connect[sheet1.cell_value(row, col)] = sheet1.cell_value(row, col + 3)

    for i in connect:
        delay = time_diff(starttime, connect[i].strip(" APM"))  # to take care of AM/PM in some attendance sheets times
        if delay >= delay_time + 2:
            late[i] = connect[i]

    return late


def dictlist(lst1, lst2):
    d = dict(zip(lst1, lst2))
    return d

def sessiontime(endtime, book2, present):
    sheet1 = book2.sheet_by_index(0)
    x = {}
    status = []
    statustime = []
    stime = {}  # local variable for dictionary of session time of each student
    for i in present:
        col = 0
        for j in range(sheet1.nrows - 1):
            # diff = 0
            if i == sheet1.cell_value(j, col):
                status.append(sheet1.cell_value(j, col + 1))
                statustime.append(sheet1.cell_value(j, col + 3))

        if len(status) == 1:
            diff = time_diff(statustime[0].lstrip(), endtime)
            stime[i] = diff
        elif len(status) % 2 == 0:
            diff = 0
            for k in range(len(status) - 1):
                if k % 2 == 0:
                    diff = diff + time_diff(statustime[k].lstrip(), statustime[k + 1].lstrip())
            # diff = diff + time_diff(statustime[k-1].lstrip(), endtime)
            stime[i] = diff
        else:
            diff = 0
            for k in range(len(status) - 1):
                if k % 2 == 0:
                    diff = diff + time_diff(statustime[k].lstrip(), statustime[k + 1].lstrip())
            diff = diff + time_diff(statustime[len(status) - 1].lstrip(), endtime)
            stime[i] = diff

        status = []
        statustime = []
    return stime


def notattentive(classduration, sessionattend):
    time80percent = .78 * classduration
    na = {}
    for i in sessionattend:
        if sessionattend[i] < time80percent:
            na[i] = sessionattend[i]
    return na


def browsestudent(): 
    filename = filedialog.askopenfilename(initialdir = "/", 
                                          title = "Select a File", 
                                          filetypes = (("Excel Workbook", 
                                                        "*.xlsx*"), 
                                                       ("all files", 
                                                        "*.*")))
    filenames.insert(0,filename)
    lab10.config(text = "File Added")

def browsepresent(): 
    filename = filedialog.askopenfilename(initialdir = "/", 
                                          title = "Select a File", 
                                          filetypes = (("Excel Workbook", 
                                                        "*.xlsx*"), 
                                                       ("all files", 
                                                        "*.*"))) 
    filenames.insert(1,filename)       
    lab11.config(text = "File Added")

def browsepath(): 
    filename = filedialog.askdirectory()
    filename = filename + "/"
    filenames.insert(2,filename)
    lab12.config(text = "Folder Selected")

def generate():
    
    global delay_time
    studentsdbPath = filenames[0]
    presentlistPath = filenames[1]
    summaryreportPath = filenames[2]

    book1 = xlrd.open_workbook(f"{studentsdbPath}")  
    book2 = xlrd.open_workbook(f"{presentlistPath}")  
    studb = createStudent(book1)
    class6,class7,class8,class9, class10, class11,class12 = classdb(studb)
    num6 = len(class6)
    num7 = len(class7)
    num8 = len(class8)
    num9 = len(class9)
    num10 = len(class10)
    num11 = len(class11)
    num12 = len(class12)
    grade = int(class_var.get())
    presentdup = presentlist(book2)
    present = removedup(presentdup)
    numpresent = len(present)
    
    
    if grade == 6:
        absent = absentlist(present, class6)
        lab9.config(text = "")
        num = num6
    elif grade == 7:
        absent = absentlist(present, class7)
        lab9.config(text = "")
        num = num7
    elif grade == 8:
        absent = absentlist(present, class8)
        lab9.config(text = "")
        num = num8
    elif grade == 9:
        absent = absentlist(present, class9)
        lab9.config(text = "")
        num = num9
    elif grade == 10:
        absent = absentlist(present, class10)
        lab9.config(text = "")
        num = num10
    elif grade == 11:
        absent = absentlist(present, class11)
        lab9.config(text = "")
        num = num11
    elif grade == 12:
        absent = absentlist(present, class12)
        lab9.config(text = "")
        num = num12
    else:
        lab9.config(text = "Invalid grade/class provided.")
    
    numabsent = len(absent)
    endtime = time_var.get()
    starttime = book2.sheet_by_index(0).cell_value(0, 3).strip(" APM")
    classduration = time_diff(starttime, endtime)
    sessionattend = sessiontime(endtime, book2, present)

    nonatten = notattentive(classduration, sessionattend)  
    delay_time = int(tol_var.get())
    
    fname = sub_var.get()
    teacher = teach_var.get()
    f = open(f"{summaryreportPath}{fname} Summary {date.today()}.txt", "w")
    f.write(f"Class summary for grade {grade} {fname} class - {teacher} ({date.today()})\n\n")

    f.write("\n")
    f.write(f"Number of students present: {numpresent}\n")
    f.write(f"Number of students absent: {numabsent}\n")
    f.write("\n")
    # writing absentee list in the summary file
    f.write("Absentee list is as follows:\n\n")
    for i in absent:
        f.write(i + "\n")

    f.write("\n")

    # writing the late joinee list in the file
    f.write(f"Students joining late by at least {delay_time} minutes are as follows:\n\n")
    for i in latejoin(starttime, book2):
        msg = i + " joined at"
        f.write(f"{msg:50s}" + "--> " + latejoin(starttime, book2)[i] + "\n")

    f.write("\n")

    # writing the % of students late for the class
    f.write("=" * 52)
    f.write("\n")
    f.write(f"Percentage of students late for the class - {(len(latejoin(starttime, book2)) / num) * 100:.2f} % |")
    f.write("\n")
    f.write("=" * 52)
    f.write("\n\n")

    # writing to the file, list of non attentive students
    f.write("Following are the students who were there in the class for less than 80%\
     of the total class duration:\n\n")

    for i in nonatten:
        msg = i + " attended the session for"
        f.write(f"{msg:60s}" + "--> " + f"{nonatten[i]} minutes\n")

    # writing the % of students late for the class
    f.write("\n")
    f.write("=" * 52)
    f.write("\n")
    f.write(f"Percentage of non-attentive students - {(len(nonatten) / num) * 100:.2f} %     |")
    f.write("\n")
    f.write("=" * 52)
    f.write("\n\n")

    f.write("Note: 0 minutes suggests that student only connected for a few seconds.")
    f.close()
    root.quit()
    root.destroy()
    webbrowser.open(f"{summaryreportPath}\{fname} Summary {date.today()}.txt")


filenames = []  
delay_time = 0
root = Tk()
root.config(bg = "#FFFFFF")
root.title("Attendance Calculator")
width = root.winfo_screenwidth()
height = root.winfo_screenheight()
root.geometry("%dx%d" % (width//2, height//2))

class_var = StringVar()
time_var = StringVar()
tol_var = StringVar()
sub_var = StringVar()
teach_var = StringVar()

lab1 = Label(root,text = "Please mention the class for which you want the class summary -",bg = "#FFFFFF",fg = "#000000",font = ("Century Gothic",12))
lab2 = Label(root,text = "At what time did your session and (HH:MM:SS)-",bg = "#FFFFFF",fg = "#000000",font = ("Century Gothic",12))
lab3 = Label(root,text = "Mention the tolerable delay in join time by students in minutes -",bg = "#FFFFFF",fg = "#000000",font = ("Century Gothic",12))
lab4 = Label(root,text = "Please enter the name of the subject for the summary file -",bg = "#FFFFFF",fg = "#000000",font = ("Century Gothic",12))
lab5 = Label(root,text = "Please enter the name of the teacher -",bg = "#FFFFFF",fg = "#000000",font = ("Century Gothic",12))
lab6 = Label(root,text = "Select the Student Database File -",bg = "#FFFFFF",fg = "#000000",font = ("Century Gothic",12))
lab7 = Label(root,text = "Select the Present List File -",bg = "#FFFFFF",fg = "#000000",font = ("Century Gothic",12))
lab8 = Label(root,text = "Select the Folder where you want to save the report -",bg = "#FFFFFF",fg = "#000000",font = ("Century Gothic",12))
lab9 = Label(root,bg = "#FFFFFF",fg = "#000000",font = ("Century Gothic",12))
lab10 = Label(root,bg = "#FFFFFF",fg = "#000000",font = ("Century Gothic",12))
lab11 = Label(root,bg = "#FFFFFF",fg = "#000000",font = ("Century Gothic",12))
lab12 = Label(root,bg = "#FFFFFF",fg = "#000000",font = ("Century Gothic",12))


ent1 = Entry(root,textvariable = class_var,bg = "#FFFFFF",relief = GROOVE,fg = "#011627",font = ("Century Gothic",12))
ent2 = Entry(root,textvariable = time_var,bg = "#FFFFFF",relief = GROOVE,fg = "#011627",font = ("Century Gothic",12))
ent3 = Entry(root,textvariable = tol_var,bg = "#FFFFFF",relief = GROOVE,fg = "#011627",font = ("Century Gothic",12))
ent4 = Entry(root,textvariable = sub_var,bg = "#FFFFFF",relief = GROOVE,fg = "#011627",font = ("Century Gothic",12))
ent5 = Entry(root,textvariable = teach_var,bg = "#FFFFFF",relief = GROOVE,fg = "#011627",font = ("Century Gothic",12))

but1 = Button(root,text = "Select",relief = FLAT,command = browsestudent,bg = "#FFFFFF",fg = "#011627",font = ("Century Gothic",12))
but2 = Button(root,text = "Select",relief = FLAT,command = browsepresent,bg = "#FFFFFF",fg = "#011627",font = ("Century Gothic",12))
but3 = Button(root,text = "Select",relief = FLAT,command = browsepath,bg = "#FFFFFF",fg = "#011627",font = ("Century Gothic",12))
but4 = Button(root,text = "Generate",relief = FLAT,command = generate,bg = "#FFFFFF",fg = "#011627",font = ("Century Gothic",12))

ent1.insert(0,"Class")
ent2.insert(0,"HH:MM:SS")
ent3.insert(0,"Tolerable Time")
ent4.insert(0,"Subject")
ent5.insert(0,"Teacher")
lab1.place(x=10,y=5)
ent1.place(x=660,y=8)
lab9.place(x=900,y=5)
lab2.place(x=10,y=55)
ent2.place(x=660,y=58)
lab3.place(x=10,y=105)
ent3.place(x=660,y=108)
lab4.place(x=10,y=155)
ent4.place(x=660,y=158)
lab5.place(x=10,y=205)
ent5.place(x=660,y=208)
lab6.place(x=10,y=255)
but1.place(x=660,y=258)
lab10.place(x=900,y=255)
lab7.place(x=10,y=305)
but2.place(x=660,y=308)
lab11.place(x=900,y=305)
lab8.place(x=10,y=355)
but3.place(x=660,y=358)
lab12.place(x=900,y=355)
but4.place(x=500,y=408)
root.mainloop()
