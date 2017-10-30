import win32com.client as win32

import comtypes, comtypes.client
import time


def startswithnumeric(strline):
    if " " in strline:
        lines=strline.split(" ")
        if lines[0].isnumeric():
            return True
        else:
            return False
    elif strline.isnumeric():
        return True
    else:
        return False

class VBA_Assistant():
    def __init__(self,**kwargs):
        self.Application_Name=kwargs["Application_Name"]
        self.File_Name=kwargs["File_Name"]
        self.Add_Line_Number=kwargs["Add_Line_Number"]
        self.Add_Error_Handler=kwargs["Add_Error_Handler"]
        self.All_Modules=kwargs["All_Modules"]
        self.Module_Name=kwargs["Module_Name"]

    
    def get_modules(self):
        mods=[]
        if self.Application_Name=="Excel":
            app = win32.gencache.EnsureDispatch('Excel.Application')
            app.Visible = False
            objfile = app.Workbooks.Open(self.File_Name)
            for m in objfile.VBProject.VBComponents:
                mods=mods + [m.Name]
            app.Quit()
        elif self.Application_Name=="Access":
            app = win32.Dispatch('Access.Application')
            app.Visible = False
            objfile = app.OpenCurrentDatabase(self.File_Name)
            for i in range(app.Modules.Count):
                mods=mods + [app.Modules(i).Name]
            app.Quit()
        return mods

    def code_modify(self):
        eh_added=False
        procname=""
        modulename =""
        n=0
        line_num=1
        skipline=False
        mods=[]
        if self.Application_Name=="Excel":
            app = win32.gencache.EnsureDispatch('Excel.Application')
            app.Visible = False
            objfile = app.Workbooks.Open(self.File_Name)
            if self.All_Modules==True:
                for m in objfile.VBProject.VBComponents:
                    mods=mods + [m.Name] 
            else:
                mods=mods + [self.Module_Name]
        elif self.Application_Name=="Access":
            app = win32.Dispatch('Access.Application')
            app.Visible = True
            objfile = app.OpenCurrentDatabase(self.File_Name.replace("/","\\"),True)
            if self.All_Modules==True:
                for m in app.Modules:
                    mods=mods + [m.Name] 
            else:
                mods=mods + [self.Module_Name]

        for m  in mods:
            skipline=False
            if self.Application_Name=='Excel':
                mod=obj.VBProject.VBComponents(m).CodeModule
                modulename=m
            elif self.Application_Name=="Access":
                mod=app.Modules(m)
                modulename=m
            # vbext_ct_StdModule
            linecount=mod.CountOfLines
            code_lines=str(mod.Lines(1,linecount)).split("\r\n")
            mod.DeleteLines(1,linecount)

            for l in code_lines:
                if "On Error GoTo "  in l:
                    eh_added=True

            for l  in code_lines:
                l1=l.lstrip()
                if l1.startswith("Public ") or l1.startswith("Private ") or l1.startswith("Sub ") or l1.startswith("End Sub") or l1.startswith("End Function") or l1.startswith("Function ") or l1.startswith("'") or l1.startswith("Option") or l1.startswith("Dim") or l1.startswith("Static") or l1.startswith("#")  or l1=="" or skipline==True or startswithnumeric(l1)==True or self.Add_Line_Number==False:
                    print(startswithnumeric(l1))
                    if l1.startswith("Sub ") or l1.startswith("Public Sub ") or l1.startswith("Private Sub ") or l1.startswith("Function ") or l1.startswith("Public Function ") or l1.startswith("Private Sub "):
                        procname=l1.replace("Public","").replace("Private","").replace("Sub","").replace("Function","").replace(" ","")
                        if eh_added==False and self.Add_Error_Handler==True:
                            mod.InsertLines(line_num,l)
                            line_num=line_num+1
                            mod.InsertLines(line_num,"On Error GoTo CatchAllError")
                        else:
                            mod.InsertLines(line_num,l)
                    elif l1.startswith("End Sub") or l1.startswith("End Function"):
                        if eh_added==False and self.Add_Error_Handler==True:
                            mod.InsertLines(line_num,"CatchAllError:")
                            line_num=line_num+1
                            mod.InsertLines(line_num,"If err.number<>0 then")
                            line_num=line_num+1
                            mod.InsertLines(line_num,'msgbox "Error at Module - ' + modulename + ', Procedure - ' + procname + ', Line - " & Erl & vbnewline & err.description')
                            line_num=line_num+1
                            mod.InsertLines(line_num,"End If")
                            line_num=line_num+1
                            mod.InsertLines(line_num,l)
                        else:
                            mod.InsertLines(line_num,l)
                    else:
                        mod.InsertLines(line_num,l)
                    print(l1)
                    skipline=False        
                else:
                    print("{0:0=2d}".format(n) + " " + l)
                    mod.InsertLines(line_num,"{0:0=2d}".format(n) + " " + l)
                    n=n+10
                if l.endswith(" _"):
                    skipline=True
                line_num=line_num+1
            if self.Application_Name=="Access":
                app.DoCmd.Save                



