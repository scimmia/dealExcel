from tkinter import *
from tkinter.filedialog import askdirectory
from tkinter.filedialog import askopenfilename


def selectPath():
    # path_ = askdirectory()
    path_ = askopenfilename()
    print(path.get())
    path.set(path_)
    print(path.get())

def doIt():
    t.insert('end', 'sdfsdf\n\n')
root = Tk()
path = StringVar()
t = Text(root)
t.grid(row=1,columnspan = 4)
Button(root, text="路径选择", command=selectPath).grid(row=0, column=0)
Entry(root, textvariable=path).grid(row=0, column=2)
Button(root, text="DoIt", command=doIt).grid(row=0, column=3)
root.mainloop()
