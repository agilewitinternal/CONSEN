from tkinter import *
#import tkinter.messagebox as tm

username = []
password = []
server_name= []
port_num = []
emailsubject = []
root = Tk()
root.geometry('500x300')
def _login_btn_clicked():
        username.append(entry_username.get())
        print(username)
        password.append(entry_password.get())
        server_name.append(entry_server.get())
        port_num.append(int(entry_port.get()))
        emailsubject.append(entry_subject.get())
        label_check.configure(text=[username,entry_password,server_name,port_num,emailsubject])
        print(emailsubject)

label_username = Label(root, text="Username")
label_password = Label(root, text="Password")
label_server = Label(root,text="Server")
label_port = Label(root,text="Port Number")
label_subject = Label(root,text="Email Subject")
label_check = Label(root,text="Check Here")

entry_username = Entry(root)
entry_password = Entry(root, show="*")
entry_server = Entry(root)
entry_port = Entry(root)
entry_subject = Entry(root)

label_username.grid(row=0, sticky=E)
label_password.grid(row=1, sticky=E)
label_server.grid(row=2,sticky=E)
label_port.grid(row=3,sticky=E)
label_subject.grid(row=4,sticky=E)
label_check.grid(row=7, sticky=E)
entry_username.grid(row=0, column=1)
entry_password.grid(row=1, column=1)
entry_server.grid(row=2,column=1)
entry_port.grid(row=3,column=1)
entry_subject.grid(row=4,column=1)



logbtn = Button(root, text="Login", command=_login_btn_clicked)
logbtn.grid(columnspan=2)


mainloop()