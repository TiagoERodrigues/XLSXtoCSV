#-*- coding: utf-8 -*-
import os.path

import pandas as pd
from tkinter import *
from tkinter import filedialog
from tkinter import messagebox
import csv

def texto():

    if len(os.listdir('C:\Conversoes')) == 0:
        messagebox.showinfo("Erro", "Arquivo não detectado")
    else:
        messagebox.showinfo("Convertido", "Arquivo convertido com sucesso")

def converter():

    texto()

    with open('C:\Conversoes\Conversao.csv', "r", encoding='utf8') as inp, open('C:\Conversoes\Planilha.txt', "w", encoding='utf8') as out:

        for file in inp:
            file = file.replace(",", ";")
            out.write(file)

    with open('C:\Conversoes\Planilha.txt', "r", encoding='utf8') as f1, open('C:\Conversoes\Conversao.csv', "w", encoding='utf8') as f2:
        f1 = pd.read_csv('C:\Conversoes\Planilha.txt')
        f1.to_csv('C:\Conversoes\Conversao.csv', encoding='utf8', index=None)
        df = pd.DataFrame(pd.read_csv('C:\Conversoes\Conversao.csv', encoding='utf8'))

    texto()
# capture file
def browseFiles():
    filename = filedialog.askopenfilename(initialdir= "/", title = "Selecione o arquivo", filetypes=(("XLSX", "*.xlsx*"),("All files", "*.*")))
    label_file_explorer.configure(text="Arquivo aberto: " + filename)
    read_file = pd.read_excel(filename)
    read_file.to_csv('C:\Conversoes\Conversao.csv', index=None)
    df = pd.DataFrame(pd.read_csv('C:\Conversoes\Conversao.csv'))

def sair():
    exit()

if __name__ == '__main__':

    pasta = r'C:\Conversoes'
    if not os.path.exists(pasta):
        os.makedirs(pasta)
    # Create the root window
    window = Tk()

    # Set window title
    window.title('XLSX para CSV')

    # Set window size
    window.geometry("400x200")

    # Set window background color
    window.config(background="white")

    # Create a File Explorer label


    label_file_explorer = Label(window,
                                text="Selecione o arquivo",
                                width=60, height=4,
                                fg="blue")

    label_file_tigas = Label(window,
                                text="Feito por Tiago Rodrigues",
                                width=60, height=4,
                                fg="black")

    button_explore = Button(window,
                            text="Abrir arquivos",
                            command=browseFiles)

    button_converter = Button(window,
                            text="CSV",
                            command=converter)

    button_sair = Button(window,
                            text="Sair",
                            command=sair)

    label_file_explorer.grid(column=0, row=1)
    label_file_tigas.grid(column=0, row=10)
    button_explore.grid(column=0, row=2)
    button_converter.grid(column=0, row=7)
    button_sair.grid(column=0, row=9)
    # main
    window.mainloop()
