import tkinter  as tk
from tkinter.filedialog import askopenfilename
from VBA_Assistant import *

def Browse():
    try:
        if appvar.get()=='Excel':
            fname = askopenfilename(filetypes=[("Excel files", "*.xlsm;*.xls;*.xlsb")])
            txt['state']='normal'
            txt.delete(0,tk.END)
            txt.insert(0,fname)
            txt['state']='readonly'
        elif appvar.get()=='Access':
            fname = askopenfilename(filetypes=[("Access files", "*.accdb;*.mdb")])
            txt['state']='normal'
            txt.delete(0,tk.END)
            txt.insert(0,fname)
            txt['state']='readonly'
    except:
        return

    if txt.get()!="":
        btn1['state']='normal'
        vb=VBA_Assistant(Application_Name=appvar.get(),File_Name=fname,Add_Line_Number=False,Add_Error_Handler=False,All_Modules=True,Module_Name=None)
        m=vb.get_modules()
        for c in m:
             cmb2['menu'].add_command(label=c, command=tk._setit(modvar, c))

def Execute():
    fname=txt.get()
    appname=appvar.get()
    al=ln.get()
    er_h=eh.get()
    mod=modvar.get()
    if al==False and er_h==False:
        return
    if fname=="":
       return
    if mod=="All":
        vb=VBA_Assistant(Application_Name=appname,File_Name=fname,Add_Line_Number=al,Add_Error_Handler=er_h,All_Modules=True,Module_Name=None)
        vb.code_modify()
    else:
        vb=VBA_Assistant(Application_Name=appname,File_Name=fname,Add_Line_Number=al,Add_Error_Handler=er_h,All_Modules=False,Module_Name=mod)
        vb.code_modify()

def Cancel():
    root.destroy()

root=tk.Tk()
root.minsize(500,300)

root.title("VBA Assistant")


appvar=tk.StringVar(root)
modvar=tk.StringVar(root)
eh=tk.IntVar(root)
ln=tk.IntVar(root)
r=1
c=1

appvar.set('Excel')
modvar.set('All')
choices={'Excel','Access'}
modules={'All'}

root.columnconfigure(0,minsize=10)
root.rowconfigure(0,minsize=10)
lbl1=tk.Label(root,text="Select Application",bg="Black",fg="White")
lbl1.grid(row=r,column=c,columnspan=2,sticky=tk.E+tk.W)

cmb1=tk.OptionMenu(root,appvar,*choices)
cmb1.grid(row=r+1,column=c,columnspan=2,sticky=tk.E+tk.W)
root.rowconfigure(r+2,minsize=10)

lbl2=tk.Label(root,text="Select File",bg="Black",fg="White")
lbl2.grid(row=r+3,columnspan=7,column=c+0,sticky=tk.E+tk.W)

txt=tk.Entry(root,state="readonly",width=80)
txt.grid(row=r+4,column=c+0,columnspan=7)

btn=tk.Button(root,text="Browse",command=Browse)
btn.grid(row=r+4,column=c+8)
root.rowconfigure(r+5,minsize=10)

lbl3=tk.Label(root,text="Select Module",bg="Black",fg="White")
lbl3.grid(row=r+6,columnspan=2,column=c+0,sticky=tk.E+tk.W)
cmb2=tk.OptionMenu(root,modvar,*modules)
cmb2.grid(row=r+7,column=c,columnspan=2,sticky=tk.E+tk.W)
root.rowconfigure(r+8,minsize=10)

c1=tk.Checkbutton(root,text="Add Error Handler",variable=eh)
c1.grid(row=r+10,column=c+0,columnspan=2,sticky=tk.W)

c2=tk.Checkbutton(root,text="Add Line Number",variable=ln)
c2.grid(row=r+11,column=c+0,columnspan=2,sticky=tk.W)

btn1=tk.Button(root,text="Execute",state=tk.DISABLED,command=Execute)
btn1.grid(row=r+12,column=c+0,sticky=tk.E+tk.W)

btn2=tk.Button(root,text="Cancel",command=Cancel)
btn2.grid(row=r+12,column=c+2,sticky=tk.E+tk.W)

root.mainloop()

