from tkinter import *

root = Tk()

a = Label(root, text = "Usuário").grid(row = 0,column = 0)
b = Label(root, text = "Senha").grid(row = 1,column = 0)
c = Label(root, text = "Partição").grid(row = 2,column = 0)
user = Entry(root)
password = Entry(root, show = "*")
partition = Entry(root)
user.grid(row=0, column=1)
password.grid(row=1, column=1)
partition.grid(row=2, column=1)

def submit():
    root.quit()

btn = Button(root, text="Submeter", command=submit)
btn.place(x=50, y=70)

root.title("Rotina do ATU")
root.geometry('350x200')
root.mainloop()
