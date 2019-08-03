from tkinter import *
import random
import time
import datetime

janelaPrincipal = Tk()
janelaPrincipal.geometry("300x300+300+200")
janelaPrincipal.title("Escala de Serviço")
janelaPrincipal.configure(background='#707070')

janelaTop = Frame(janelaPrincipal, width=300, height=50, bd=4, relief='raise', background='red')
janelaTop.pack(side=TOP)

janelaEsquerda = Frame(janelaPrincipal, width=150, height=50, bd=4, relief='raise', background='blue')
janelaEsquerda.pack(side=LEFT)

janelaDireita = Frame(janelaPrincipal, width=150, height=50, bd=4, relief='raise', background='black')
janelaDireita.pack(side=RIGHT)

textoTitulo = Label(janelaTop, font=('arial', 20), text="ESCALA DE SERVIÇO", background='white')
textoTitulo.grid(row=0, column=0)

janelaPrincipal.mainloop()


