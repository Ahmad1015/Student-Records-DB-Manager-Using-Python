import tkinter
from tkinter import *

import pyodbc

msa_drivers = [x for x in pyodbc.drivers() if 'ACCESS' in x.upper()]
print(f'MS-Access Drivers : {msa_drivers}')

try:
    con_string = r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=D:\ICT projects\STUDENT.accdb;'
    conn = pyodbc.connect(con_string)
    print("Connected To Database")
except pyodbc.Error as e:
    print("Error in Connection", e)
cur = conn.cursor()

# Design the Student Database Form
root = Tk()
root.geometry("650x400")

#
Label(root, text="Student Database Form", font="Arial 12 bold", foreground='blue').grid(row=0, column=0)
global message
message = Label(root, text="Message Will Appear Here!", foreground='red')
sname = Label(root, text='Student Name', font="ar 10 bold")
fname = Label(root, text='Father Name', font="ar 10 bold")
cnic = Label(root, text='CNIC# (P.Key)', font="ar 10 bold")
search = Label(root, text='Search Record', font="ar 10 bold")
city = Label(root, text='City', font="ar 10 bold")
marks = Label(root, text='Marks', font="ar 10 bold")

message.grid(row=0, column=1)
sname.grid(row=2, column=0)
fname.grid(row=3, column=0)
cnic.grid(row=4, column=0)
search.grid(row=4, column=2)
city.grid(row=5, column=0)
marks.grid(row=6, column=0)

sNameValue = StringVar()
fNameValue = StringVar()
cnicValue = StringVar()
cityValue = StringVar()
marksValue = IntVar()

sNameEntery = Entry(root, textvariable=sNameValue, width=30, font='ar 12 bold')
fNameEntery = Entry(root, textvariable=fNameValue, width=30, font='ar 12 bold')
cnicEntery = Entry(root, textvariable=cnicValue, width=30, font='ar 12 bold')
cityEntery = Entry(root, textvariable=cityValue, width=30, font='ar 12 bold')
marksEntery = Entry(root, textvariable=marksValue, width=30, font='ar 12 bold')

sNameEntery.grid(row=2, column=1, pady=15)
fNameEntery.grid(row=3, column=1, pady=15)
cnicEntery.grid(row=4, column=1, pady=15)
cityEntery.grid(row=5, column=1, pady=15)
marksEntery.grid(row=6, column=1, pady=15)

i = 0


# FirstRecord function
def FirstRecord():  # Fine
    global message
    global i, flagfromLast
    flagfromLast = False
    cur.execute("SELECT * FROM Student ")
    data = cur.fetchall()[0]
    sName, fName, CNIC, City, marks = data
    sNameValue.set(sName)
    fNameValue.set(fName)
    cnicValue.set(CNIC)
    cityValue.set(City)
    marksValue.set(marks)
    conn.commit()
    print("First Record")
    i = 0
    message.grid_forget()
    message = Label(root, text="First Record of the Student Table!", foreground='Green', font='20')
    message.grid(row=0, column=1)


# ClearRecord function
def ClearRecord():
    global message
    sNameValue.set("")
    fNameValue.set("")
    cnicValue.set("")
    cityValue.set("")
    marksValue.set("")
    print("Record Cleared")
    message.grid_forget()
    message = Label(root, text="Records Cleared!", foreground='Green', font='20')
    message.grid(row=0, column=1)


i = 0


def NextRecord():
    global i
    i = i + 1
    global message
    global flagfromLast
    try:
        if flagfromLast:
            i = -1
        cur.execute("SELECT * FROM Student ")
        data = cur.fetchall()[i]
        sName, fName, CNIC, City, marks = data
        sNameValue.set(sName)
        fNameValue.set(fName)
        cnicValue.set(CNIC)
        cityValue.set(City)
        marksValue.set(marks)
        conn.commit()
        print("Next Record")
        message.grid_forget()
        if flagfromLast:
            message = Label(root, text="Already on Last Record!", foreground='Red', font='Helvetica 11 bold')
        else:
            message = Label(root, text="Next Record of the Student Table!", foreground='Green', font='20')
        message.grid(row=0, column=1)
    except IndexError:
        message.grid_forget()
        message = Label(root, text="On Last Record - Cannot Move Forward", foreground='Red', font='Helvetica 11 bold')
        message.grid(row=0, column=1)


# PreviousRecord function
def PreviousRecord():
    global i
    if i == 0:
        i = 0
        flag1 = True
    else:
        i = i - 1
    global message
    cur.execute("SELECT * FROM Student ")
    data = cur.fetchall()[i]
    sName, fName, CNIC, City, marks = data
    sNameValue.set(sName)
    fNameValue.set(fName)
    cnicValue.set(CNIC)
    cityValue.set(City)
    marksValue.set(marks)
    conn.commit()
    print("Previous Record")
    message.grid_forget()
    if flag1:
        message = Label(root, text="Already on last Record!", foreground='Red', font='Helvetica 11 bold')
        message.grid(row=0, column=1)
    else:
        message = Label(root, text="Previous Record of the Student Table!", foreground='Green', font='20')
        message.grid(row=0, column=1)


flagfromLast = False


# LastRecord function
def LastRecord():
    global message
    global flagfromLast
    flagfromLast = True
    cur.execute("SELECT * FROM Student ")
    data = cur.fetchall()[-1]
    sName, fName, CNIC, City, marks = data
    sNameValue.set(sName)
    fNameValue.set(fName)
    cnicValue.set(CNIC)
    cityValue.set(City)
    marksValue.set(marks)
    conn.commit()
    print("Last Record")
    message.grid_forget()
    message = Label(root, text="Last Record of the Student Table!", foreground='Green', font='20')
    message.grid(row=0, column=1)


# InsertRecord function
def InsertRecord():
    global message
    WrongName = False
    flag = False
    wrongCNIC = False
    try:
        Sname = sNameValue.get()
        fname = fNameValue.get()
        CNIC = cnicValue.get()
        City = cityValue.get()
        marks = marksValue.get()
        for x in range(len(CNIC)):
            slicing = CNIC[x:x + 1]
            if 48 > ord(slicing) or 57< ord(slicing):
                wrongCNIC = True
                raise ValueError
        for x in range(len(Sname)):
            slicing = Sname[x:x + 1]
            if ord(slicing) < 65 or ord(slicing) > 122:
                WrongName = True
                raise ValueError
        for x in range(len(fname)):
            slicing = fname[x:x + 1]
            if ord(slicing) < 65 or ord(slicing) > 122:
                WrongName = True
                raise ValueError
        if CNIC == '' or len(CNIC) > 13:
            if len(CNIC) > 13:
                flag = True
                raise TypeError("CNIC too long")
            else:
                raise TypeError("CNIC Required")
        if marks < 0 or marks > 100:
            raise ValueError
        cur.execute(f"INSERT INTO Student (Sname,Fname,CNIC,City,marks) Values (?,?,?,?,?);",
                    (Sname, fname, CNIC, City, marks))
        conn.commit()
        print("Insert Record")
        message.grid_forget()
        message = Label(root, text="Successfully inserted a Record into the Student Table!", foreground='Green',
                        font='20')
        message.grid(row=0, column=1)
    except TypeError:
        message.grid_forget()
        if flag:
            message = Label(root, text="CNIC Too long", foreground='Red', font='Helvetica 11 bold')
        else:
            message = Label(root, text="CNIC cannot be left Empty", foreground='Red', font='Helvetica 11 bold')
        message.grid(row=0, column=1)
    except pyodbc.Error as e:
        message.grid_forget()
        message = Label(root, text="CNIC already exists", foreground='Red', font='Helvetica 11 bold')
        message.grid(row=0, column=1)
    except ValueError:
        message.grid_forget()
        if WrongName:
            message = Label(root, text="Name cannot have numbers or Special Characters", foreground='Red',
                            font='Helvetica 11 bold')
        else:
            message = Label(root, text="Marks  should be between zero and 100 inclusive", foreground='Red',
                            font='Helvetica 11 bold')
        if wrongCNIC:
            message = Label(root, text="CNIC cannot have numbers or Special Characters", foreground='Red',
                            font='Helvetica 11 bold')
        message.grid(row=0, column=1)


# UpdateRecord function
def UpdateRecord():
    global message
    try:
        CNIC = cnicValue.get()
        if CNIC == '':
            raise TypeError("CNIC Required")
        cur.execute("select* from Student where CNIC=?", (CNIC))
        data = cur.fetchall()[0]
        Sname = sNameValue.get()
        Fname = fNameValue.get()
        City = cityValue.get()
        marks = marksValue.get()
        cur.execute("UPDATE Student set Sname=?,Fname=?,City=?,marks=?  WHERE CNIC=? ",
                    (Sname, Fname, City, marks, CNIC))
        conn.commit()
        print("Update Record")
        message.grid_forget()
        message = Label(root, text="Successfully Updated Record of the Student Table!", foreground='Green', font='20')
        message.grid(row=0, column=1)
    except IndexError:
        message.grid_forget()
        message = Label(root, text="Please enter Correct CNIC to update", foreground='Red', font='Helvetica 11 bold')
        message.grid(row=0, column=1)
    except TypeError:
        message.grid_forget()
        message = Label(root, text="CNIC cannot be left Empty", foreground='Red', font='Helvetica 11 bold')
        message.grid(row=0, column=1)


# DeleteRecord function
def DeleteRecord():  # Fine
    global message
    CCnic = cnicValue.get()
    try:
        if CCnic == '':
            raise TypeError("CNIC Required")
        cur.execute("select* from Student where CNIC=?", (CCnic))
        data = cur.fetchall()[0]
        cur.execute("DELETE FROM Student where CNIC=?", (CCnic))
        conn.commit()
        print("Deleted Record")
        message.grid_forget()
        message = Label(root, text="Deleted Record of the Student Table!", foreground='Green', font='20')
        message.grid(row=0, column=1)
    except IndexError:
        message.grid_forget()
        message = Label(root, text="No Such Record Found", foreground='Red', font='Helvetica 11 bold')
        message.grid(row=0, column=1)
    except TypeError:
        message.grid_forget()
        message = Label(root, text="CNIC cannot be left Empty", foreground='Red', font='Helvetica 11 bold')
        message.grid(row=0, column=1)


# SearchRecord function
def SearchRecord():
    global message
    CCnic = cnicValue.get()
    try:
        if CCnic == '':
            raise TypeError("CNIC Required")
        cur.execute("select* from Student where CNIC=?", (CCnic))
        data = cur.fetchall()[0]
        sName, fName, CNIC, City, marks = data
        sNameValue.set(sName)
        fNameValue.set(fName)
        cnicValue.set(CNIC)
        cityValue.set(City)
        marksValue.set(marks)
        conn.commit()
        print("Search Record Found")
        message.grid_forget()
        message = Label(root, text="Search Record from the Student Table Found!", foreground='Green',
                        font='Helvetica 11 bold')
        message.grid(row=0, column=1)
    except IndexError:
        message.grid_forget()
        message = Label(root, text="No Record Found against CNIC", foreground='Red', font='Helvetica 11 bold')
        message.grid(row=0, column=1)
    except TypeError:
        message.grid_forget()
        message = Label(root, text="Please enter CNIC to search", foreground='Red', font='Helvetica 11 bold')
        message.grid(row=0, column=1)


Button(text="CLEAR", command=ClearRecord, background='gray', foreground='blue', font='ar 10 bold').grid(row=7, column=0)
Button(text="FIRST", command=FirstRecord, background='gray', foreground='blue', font='ar 10 bold').grid(row=7, column=1)
Button(text="NEXT", command=NextRecord, background='gray', foreground='blue', font='ar 10 bold').grid(row=7, column=2)
Button(text="PREVIOUS", command=PreviousRecord, background='gray', foreground='blue', font='ar 10 bold').grid(row=9,
                                                                                                              column=0)
Button(text="LAST", command=LastRecord, background='gray', foreground='blue', font='ar 10 bold').grid(row=9, column=1)
Button(text="INSERT", command=InsertRecord, background='gray', foreground='blue', font='ar 10 bold').grid(row=9,
                                                                                                          column=2)
Button(text="UPDATE", command=UpdateRecord, background='gray', foreground='blue', font='ar 10 bold').grid(row=11,
                                                                                                          column=0)
Button(text="DELETE", command=DeleteRecord, background='gray', foreground='blue', font='ar 10 bold').grid(row=11,
                                                                                                          column=1)
Button(text="SEARCH", command=SearchRecord, background='gray', foreground='blue', font='ar 10 bold').grid(row=11,
                                                                                                          column=2)

root.mainloop()
