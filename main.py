import openpyxl 
from openpyxl import *
from openpyxl import load_workbook
from tkinter import *
from tkinter import ttk
from tkinter import simpledialog
import tkinter as tk
import os
from datetime import datetime
from tkinter import PhotoImage 

root = tk.Tk()
root.title("PyXLSXManager")
root.minsize(300, 300)

#Déclaration variable global
wb = load_workbook('hello.xlsx')
ws = wb.active
max_column_plus_one = ws.max_column + 1
max_row_plus_one = ws.max_row + 1
numero_ref=""
date_arrive=""
date_fin=""
status=""
responsable=""
description=""
technologie=""

#Déclaration Fonction
def quitter_programme():
    root.destroy()
def ecriture_excel():
	global wb,ws,max_row_plus_one,numero_ref,date_arrive,date_arrive,date_fin,status,responsable,description,technologie
	ws.cell(row=max_row_plus_one,column=1,value=numero_ref)
	ws.cell(row=ws.max_row,column=2,value=date_arrive)
	ws.cell(row=ws.max_row,column=3,value=date_fin)
	ws.cell(row=ws.max_row,column=4,value=status)
	ws.cell(row=ws.max_row,column=4,value=responsable)
	ws.cell(row=ws.max_row,column=5,value=description)
	ws.cell(row=ws.max_row,column=6,value=technologie)

def afficher_data():
	global ws,root
	top_fenetre=Toplevel(root)
	top_fenetre.minsize(300, 250)
	my_scrollbar = Scrollbar(top_fenetre, orient=VERTICAL)
	my_listbox = Listbox(top_fenetre, width=100, yscrollcommand=my_scrollbar.set, selectmode=NONE)
	my_listbox.config
	my_scrollbar.pack(side=RIGHT, fill=Y)
	my_listbox.pack(pady=15)
	my_list = []
	for value in ws.iter_rows(min_row=1, max_row=11, min_col=1, max_col=8,values_only=True):
		my_list.append(value)
	dataIndex = 0
	for item in my_list:
		my_listbox.insert(END, str(item))
		dataIndex = dataIndex + 1




#Grid
root.columnconfigure((0),weight =1,uniform='a')
root.rowconfigure((0,1,2,3,4,5,6,7),weight =1,uniform='a')

label_bienvenue = Label(master= root, text="Bienvenue dans PyXLSXManager ", font="Helvetica 18 bold")
label_bienvenue.grid(row = 0,column = 0, sticky="n")

image_main = PhotoImage(file="png/main_image.png")
label_image = Label(master = root, image= image_main)
label_image.grid(row = 1,column = 0,sticky="nwse")

bouton_ajouter = Button(master= root, text="Ajoutez les informations", font="Helvetica 12"  , width = 25)
bouton_ajouter.grid(row = 3,column = 0, sticky="n")

bouton_voir = Button(master= root, text="Voir les 10 dernière entrées",font="Helvetica 12",command=afficher_data , width = 25)
bouton_voir.grid(row = 4,column = 0, sticky="n")

bouton_quitter = Button(root, text="Quitter",font="Helvetica 12" ,command=quitter_programme)
bouton_quitter.grid(row = 7,column = 0, sticky="n")



# Save the file
wb.save("hello.xlsx")
#Mainloop
root.mainloop()