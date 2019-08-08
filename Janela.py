from tkinter import *


def soma():
    if str(valor1.get()).isnumeric() and str(valor2.get()).isnumeric():
        num1 = int(valor1.get())
        num2 = int(valor2.get())
        resultado = num1+num2
    else:
        resultado = "ERRO: Dados inválidos!"
    lbResultado["text"] = resultado
def subrair():
    if str(valor1.get()).isnumeric() and str(valor2.get()).isnumeric():
        num1 = int(valor1.get())
        num2 = int(valor2.get())
        resultado = num1-num2
    else:
        resultado = "ERRO: Dados inválidos!"
    lbResultado["text"] = resultado


def fechar():
    janelaPrincipal.destroy()


janelaPrincipal = Tk()
janelaPrincipal.geometry("300x300+500+300")
janelaPrincipal.title("Escala de Serviço")

valor1 = Entry(janelaPrincipal, width=15, bg="white")
valor2 = Entry(janelaPrincipal, width=15, bg="white")
botaoSoma = Button(janelaPrincipal, text="SOMA", width=10, command=soma)
botaoSubtrair = Button(janelaPrincipal, text='SUBTRAIR', width=10, command=subrair)
lbResultado = Label(janelaPrincipal, text="Resultado")
botaoFechar = Button(janelaPrincipal, text="Fechar", command=fechar)

valor1.grid(row=0, column=0)
valor2.grid(row=0, column=1)
botaoSoma.grid(row=1, column=0, sticky=E)
botaoSubtrair.grid(row=1, column=1, sticky=W)
lbResultado.grid(row=2, column=0, columnspan=2)
botaoFechar.grid(row=3, column=0, columnspan=2, sticky=E+W)



janelaPrincipal.mainloop()


