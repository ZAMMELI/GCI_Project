from tkinter import *
from tkinter.ttk import *
from datetime import datetime
from time import strftime

from PIL import Image,ImageTk


##################################
# this func to transforming files
# to excel document
##################################

import os
import sys
import glob
import xlsxwriter

###### in determinate mode ######

# creating tkinter window for intro

intro = Tk()
can = Canvas(intro,width=650,height=500,bg="white")
photo=ImageTk.PhotoImage(Image.open("K:\\Agilent_ICT\\Boards\\GCI\\icons\\ICT.jpg"))
can.create_image(80,80,image=photo,anchor=NW)
can.grid(row=1,rowspan=3,padx=10,pady=5)
intro.title("INTRO")
intro.after(1000,lambda:intro.destroy())

# creating tkinter window app
root = Tk()
root.title('GCI_APP_Developed by Aymen ZAMMELI')
root.geometry('450x320')
root.config(bg="#161164")

# this func is used to
# add exit button and quit app
def abFunc():
        msgAbout()

def msgAbout():
        txt = Toplevel(root)
        txt.title('about application')
        txt.geometry('550x500')
        #txt.config(bg="#552fdf")

        ablbl = Label(txt, text="The prior sections introduced sinusoidal sources, phasors, and the phasor representation of some")
        ablbl.pack(anchor = 'center')

my_menu=Menu(root)
root.config(menu=my_menu)
my_menu.add_command(label='About',command= abFunc)

# This function is used to
# display time on the label
def time():
	string = strftime('%H:%M:%S %p')
	lbl.config(text = string)
	lbl.after(1000, time)

# Styling the label widget so that clock
# will look more attractive
lbl = Label(root, font = ('calibri', 25, 'bold'),
			background = "#161164",
			foreground = 'white')
# Placing clock at the center
# of the tkinter window
lbl.pack(anchor = 'center')
time()
################# name app
#app_lbl = Label(root, font = ('calibri', 15, 'bold'),background = 'white', foreground = 'black')
app_lbl = Label(root,text = 'ASTEELFLASH TUNISIA', font = ('calibri', 15, 'bold'),background = 'white', foreground = 'black')
# Placing name at the center
# of the tkinter window
app_lbl.pack(anchor = 'center')

#########################################148-157-105

def Main_GCI():
        path = "K:\\Agilent_ICT\\Boards\\NomDesc\\*"

        width = len("this is the meduim size")      
        lis= []  

        # Cretae a xlsx file
        xlsx_File = xlsxwriter.Workbook('Nombre_de_descentes_totales.xlsx')
        # Add new worksheet
        sheet_names = xlsx_File.add_worksheet()
        #format objects
        cfg = xlsx_File.add_format({'bold': True, 'font_color': 'green','num_format':'#####'})
        cfr = xlsx_File.add_format({'bold': True, 'font_color': 'red','num_format':'#####'})
        af = xlsx_File.add_format({'bold': True, 'font_color': 'blue','num_format':'#####'})
        gmao = xlsx_File.add_format({'bold': True, 'font_color': 'orange','num_format':'#####'})
        form = xlsx_File.add_format({'bold': True, 'font_color': 'purple','num_format':'#####'})

        row = 0
        column = 0
        sheet_names.write(row,column,"INTERFACE",form)
        sheet_names.write(row,column+1,"ID_GMAO",form)
        sheet_names.write(row,column+2,"AUTOFILE",form)
        sheet_names.write(row,column+3,"COUNTER",form)
        sheet_names.write(row,column+4,"T_TEST",form)
        
        for nom in glob.glob(path):
            #print(nom)
            lis.append(nom)
        print(lis)
        r = len(lis)
        for i in range(r):        
            with open(lis[i] + "\\cpteur_GCI_save.txt", 'r')as infile:
                print(lis[i])
                l = infile.readline()
                if l != "":
                    sheet_names.write(row+1,column, l)
                    #print('si 1')
                l = infile.readline()
                if l != "":
                    sheet_names.write(row+1,column+1, l, gmao)
                    #print('si 1')
                l = infile.readline()
                if l != "":
                    sheet_names.write(row+1,column+2, l, af)
                    #print('si 2')
                l = infile.readline()
                
                if l != "":
                        c = float(l)
                        if c >= 60000:
                                sheet_names.write(row+1,column+3, l, cfr)
                        else:
                                sheet_names.write(row+1,column+3, l, cfg)
                        
                        #print('si 3')
                    #print(l)
                l = infile.readline()          
                if l != "":
                    sheet_names.write(row+1,column+4, l)
                    #print('si 1')
                row += 1          

        xlsx_File.close()
                    
        print("the length of folder is : " + str(len(lis)))
        print('**** work done ****')
        
        #auto-launch excel file
        os.system("start EXCEL.EXE Nombre_de_descentes_totales.xlsx")

##########################################

# Progress bar widget
progress = Progressbar(root, orient = HORIZONTAL,
			length = 350, mode = 'determinate')

# Function responsible for the updation
# of the progress bar value
def bar():
	import time
	progress['value'] = 20
	root.update_idletasks()
	time.sleep(0.11)

	progress['value'] = 40
	root.update_idletasks()
	time.sleep(0.11)

	progress['value'] = 50
	root.update_idletasks()
	time.sleep(0.11)

	progress['value'] = 60
	root.update_idletasks()
	time.sleep(0.11)

	progress['value'] = 80
	root.update_idletasks()
	time.sleep(0.11)
	progress['value'] = 100

	Main_GCI()
     
progress.pack(pady = 10)

################################################
#    test for max compter to do preventif      #
################################################


variable1 = StringVar(root)
variable2 = StringVar(root)


# reset function
def RAZ():
        messagebox()
                    

def messagebox():

        #function to get autofile and change counter file to empty
        def getInput():
                inp = inputtxt.get(1.0, "end-1c")
                print("inp :" +inp)
                folder = "K:\\Agilent_ICT\\Boards\\Nombre_de_descentes_totales\\" + str(inp)
                print("path :" +folder)
                
                with open(folder + "\\cpteur", 'r')as infile:
                        xdata = infile.read()
                        print(type(xdata))
                        print("read autofile :" +xdata)
                        infile.close()

                ndata = xdata.replace(xdata,"0")
                print("replace autofile :" +ndata)
                
                with open(folder + "\\cpteur", 'w')as infile:
                        infile.write(ndata)
                        print("write autofile :" +ndata)
                        
                        infile.close()
                toplevel.destroy()
                
        #second message box creation
        toplevel = Toplevel(root)
        toplevel.title("remise à zero de compteur interface")
        toplevel.geometry("300x150")
        #toplevel.config(bg="#5FB691")


        #create label to tret text
        L0 = Label(toplevel, text="you will set autofile to zero !",font = ('calibri', 12, 'bold'),foreground = 'red')
        L0.grid(row=0,column=1)
        #create label to tret text
        L1 = Label(toplevel, text="autofile :")
        L1.grid(row=1)

        
        #create text box
        inputtxt = Text(toplevel,height=1,width=12)
        inputtxt.grid(row=1,column=1)

        
        #create buttons of msg box
        b1=Button(toplevel,text="OK",command=getInput,width = 15)
        b1.grid(row=2, column=1)
                
        
        b2=Button(toplevel,text="CANCEL",command=toplevel.destroy,width = 15)
        b2.grid(row=3, column=1)

        
       
            
#******************************************************************************
#******************************************************************************


Button(root, text = 'Start', command = bar).pack(pady = 10)
Button(root, text="RAZ", command=RAZ).pack(pady = 10)
Button(root, text = 'Exit', command = root.destroy).pack(pady = 10)
# infinite loop
root.iconbitmap(r'K:\Agilent_ICT\Boards\GCI\icons\icon.ico')
root.mainloop()


