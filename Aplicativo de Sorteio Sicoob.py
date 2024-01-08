import tkinter as tk
from tkinter import *
from PIL import Image, ImageTk
import pandas as pd
import random
import os
import time
import threading
import pygame

lista = os.getcwd().split("\\")
path_user = ""
for item in os.path.expanduser("~").split("\\"):
    path_user += item + '/'
    if 'Automatização' in item:
        break
bateria = path_user+"OneDrive - Sicoob/Documentos - Sicoob UniRondônia/27. Atualizações Python/Sorteio Sicoob/bateria.mp3"
prato = path_user+"OneDrive - Sicoob/Documentos - Sicoob UniRondônia/27. Atualizações Python/Sorteio Sicoob/prato.mp3"
def botao1_clicado():
    # Cria uma nova thread para realizar o sorteio
    thread_sorteio = threading.Thread(target=realizar_sorteio)
    botao.place_forget()  # Torna o botão atual invisível
    thread_sorteio.start()

def esconder():
    botão_sortear.place_forget()
    segundo_botao_clicado.set(False)

def botao2_clicado():
    segundo_botao_clicado.set(True)
    print("Segundo botão foi clicado!")

def tocar_som(arquivo, loop=True):
    pygame.mixer.init()
    pygame.mixer.music.load(arquivo)
    pygame.mixer.music.play(-1 if loop else 0)

def parar_som():
    pygame.mixer.music.stop()


def realizar_sorteio():
    print("Iniciando sorteio...")
    time.sleep(1)
    label_numero= tk.Label(root, text="", font=("Asap", 45,"bold"),fg="#FFFFFF", justify="center", anchor="center", bg="#00A091")
    label_pa = tk.Label(root, text="", font=("Asap", 35),fg="#FFFFFF", justify="center", anchor="center", bg="#00A091")
    label_nome = tk.Label(root, text="", font=("Asap", 35),fg="#FFFFFF", justify="center", anchor="center", bg="#00A091")
    def limpar_tela():
        label_numero.destroy()
        label_nome.destroy()
        label_pa.destroy()
        botao.place(x=1171, y=579)  # Mostra o botão "INICIAR" novamente
        botão_reeniciar.place_forget() 
    
    botão_reeniciar = tk.Button(root, text="VOLTAR", command=limpar_tela, bg=cor_hexadecimal, fg="white", font=("Asap", 32), border=5, relief=tk.GROOVE, padx=10, pady=5, activebackground="#138FAB")
    
    caminho_arquivo = path_user+"OneDrive - Sicoob/Documentos - Sicoob UniRondônia/27. Atualizações Python/Sorteio Sicoob/Cupons.xlsx"
    df = pd.read_excel(caminho_arquivo)
    codigos = df['Código'].tolist()
    
    total = str(len(codigos))

    lb_aviso = tk.Label(root, text="Total de cupons: "+ total, font=("Asap", 50,"bold"),fg="#093640", justify="center", anchor="center")
    lb_aviso.place(x=931, y=550)
    root.update()
    time.sleep(4)
    lb_aviso.config(text="Preparando o Sorteio")
    root.update()
    time.sleep(1)
    lb_aviso.config(text="Preparando o Sorteio.")
    root.update()
    time.sleep(1)
    lb_aviso.config(text="Preparando o Sorteio..")
    root.update()
    time.sleep(1)
    root.update()
    tocar_som(bateria, loop=True)

    # Sorteando um valor entre os códigos
    numero_sorte = random.choice(codigos)
    # Obtendo o nome correspondente ao código sorteado
    linha_sorteada = df[df['Código'] == numero_sorte]
    nome_sorteado = linha_sorteada['Nome do Cooperado'].values[0]
    pa_sorteado = str(linha_sorteada['Número PA Carteira'].values[0])
    #ganhadores
    df_atual = linha_sorteada
    df_anterior = pd.read_excel(path_user+"OneDrive - Sicoob/Documentos - Sicoob UniRondônia/27. Atualizações Python/Sorteio Sicoob/Ganhadores.xlsx")
    df_anterior_cleaned = df_anterior.dropna(axis=1, how='all')
    df_atual_cleaned = df_atual.dropna(axis=1, how='all')

    # Concatenar DataFrames
    ganhadores = pd.concat([df_anterior_cleaned, df_atual_cleaned])
    ganhadores.to_excel(path_user+"OneDrive - Sicoob/Documentos - Sicoob UniRondônia/27. Atualizações Python/Sorteio Sicoob/Ganhadores.xlsx", index=False)
    print("Base Atualizada")


    # Excluindo a linha correspondente ao código sorteado
    df = df[df['Código'] != numero_sorte]
    num = str(numero_sorte)
    n1 = num[0]
    n2 = num[1]
    n3 = num[2]
    n4 = num[3]
    n5 = num[4]

    # Salvando as alterações de volta no arquivo Excel
    
    df.to_excel(caminho_arquivo, index=False)
    
    # Load the GIF image
    gif1 = Image.open(path_user+"OneDrive - Sicoob/Documentos - Sicoob UniRondônia/27. Atualizações Python/Sorteio Sicoob/0.gif")
    lb_aviso.destroy()
    
    
    botão_sortear.place(x=1100, y=800) 
    root.update()

    # Create a list of frames
    frames = []
    for i in range(gif1.n_frames):
        gif1.seek(i)
        frames.append(ImageTk.PhotoImage(gif1))
    
    
    # Create a label widget to display the frames
    label_anima1 = Label(root)
    label_anima1.place(x=931, y=491)
    label_anima2 = Label(root)
    label_anima2.place(x=1058, y=491)
    label_anima3 = Label(root)
    label_anima3.place(x=1185, y=491)
    label_anima4 = Label(root)
    label_anima4.place(x=1312, y=491)
    label_anima5 = Label(root)
    label_anima5.place(x=1439, y=491)


    # Define a function to play the animation
    def play_animation1(frame_idx):
        label_anima1.config(image=frames[frame_idx])
        root.after(50, play_animation1, (frame_idx+1) % len(frames))
        root.update()
        
    def play_animation2(frame_idx):
        label_anima2.config(image=frames[frame_idx])
        root.after(50, play_animation2, (frame_idx+1) % len(frames))
        root.update()

    def play_animation3(frame_idx):
        label_anima3.config(image=frames[frame_idx])
        root.after(50, play_animation3, (frame_idx+1) % len(frames))
        root.update()
    
    def play_animation4(frame_idx):
        label_anima4.config(image=frames[frame_idx])
        root.after(50, play_animation4, (frame_idx+1) % len(frames))
        root.update()

    def play_animation5(frame_idx):
        label_anima5.config(image=frames[frame_idx])
        root.after(50, play_animation5, (frame_idx+1) % len(frames))
        root.update()
    
    

    # Start playing the animation
    play_animation1(0)
    play_animation2(3)
    play_animation3(8)
    play_animation4(6)
    play_animation5(1)

    
    largura_retangulo = 112  
    altura_retangulo = 148  
    cor_retangulo = "#00A091"  

    
    while segundo_botao_clicado.get() != True:
        print("aguardo")
    esconder()
    

    lb_n1 = tk.Label(root, text=n1, font=("Asap", 78,"bold"),fg="#FFFFFF", justify="center", anchor="center", bg="#00A091")
    canvas = tk.Canvas(root, width=largura_retangulo, height=altura_retangulo, bg=cor_retangulo)

    lb_n2 = tk.Label(root, text=n2,  font=("Asap", 78,"bold"),fg="#FFFFFF", justify="center", anchor="center", bg="#00A091")
    canvas2 = tk.Canvas(root, width=largura_retangulo, height=altura_retangulo, bg=cor_retangulo)

    lb_n3 = tk.Label(root, text=n3, font=("Asap", 78,"bold"),fg="#FFFFFF", justify="center", anchor="center", bg="#00A091")
    canvas3 = tk.Canvas(root, width=largura_retangulo, height=altura_retangulo, bg=cor_retangulo)

    lb_n4 = tk.Label(root, text=n4,  font=("Asap", 78,"bold"),fg="#FFFFFF", justify="center", anchor="center", bg="#00A091")
    canvas4 = tk.Canvas(root, width=largura_retangulo, height=altura_retangulo, bg=cor_retangulo)
    
    lb_n5 = tk.Label(root, text=n5, font=("Asap", 78,"bold"),fg="#FFFFFF", justify="center", anchor="center", bg="#00A091")
    canvas5 = tk.Canvas(root, width=largura_retangulo, height=altura_retangulo, bg=cor_retangulo)

    

    def destruir_animacao1():
        label_anima1.destroy()
       
        lb_n1.place(x=961, y=495)
        # Create the canvas
        canvas.place(x=931, y=491)
        # Raise the label above the canvas
        lb_n1.lift(canvas)
    
    def destruir_animacao2():
        label_anima2.destroy()
        # Create and place the label
        
        lb_n2.place(x=1088, y=495)
        # Create the canvas
       
        canvas2.place(x=1058, y=491)
        # Raise the label above the canvas
        lb_n2.lift(canvas2)
     

    def destruir_animacao3():
        label_anima3.destroy()
        lb_n3.place(x=1215, y=495)
        canvas3.place(x=1185, y=491)
        lb_n3.lift(canvas3)

    def destruir_animacao4():
        label_anima4.destroy()
        lb_n4.place(x=1340, y=495)
        canvas4.place(x=1312, y=491)
        lb_n4.lift(canvas4)

    def destruir_animacao5():
        label_anima5.destroy()
        lb_n5.place(x=1469, y=495)
        canvas5.place(x=1439, y=491)
        lb_n5.lift(canvas5)
    
    destruir_animacao1()
    time.sleep(1)
    destruir_animacao2()
    time.sleep(2)
    destruir_animacao3()
    time.sleep(1)
    destruir_animacao4()
    time.sleep(2)
    destruir_animacao5()
    time.sleep(1)

    
    tocar_som(prato,loop=False)

    def label_destroy(label, canvas):
        label.destroy()
        canvas.destroy()
    def numero_gif():
        root.after(0, lambda: label_destroy(lb_n1, canvas))
        root.after(0, lambda: label_destroy(lb_n2, canvas2))
        root.after(0, lambda: label_destroy(lb_n3, canvas3))
        root.after(0, lambda: label_destroy(lb_n4, canvas4))
        root.after(0, lambda: label_destroy(lb_n5, canvas5))
    numero_gif()
   
  
    meio = len(nome_sorteado) // 2
    if len(nome_sorteado) > 30:
        # Divide a string ao meio e adiciona uma quebra de linha
        metade1 = nome_sorteado[:meio]
        metade2 = nome_sorteado[meio:]
        nome_sorteado = metade1 + "\n" + metade2
    else:
        pass
    
    label_numero.config(text=numero_sorte)
    # Atualizar o texto do Label
    label_nome.config(text=nome_sorteado)
    label_pa.config(text="PA: " + pa_sorteado)
    label_numero.place(x=1100, y=400)
    label_pa.place(x=990, y=650)
    label_nome.place(x=990, y=500)

    botão_reeniciar.place(x=1100, y=800)
    root.update()
    

        
    
    

    

def sair_tela_cheia(event):
    root.attributes("-fullscreen", False)


root = tk.Tk()
root.title("Sorteios Sicoob Unirondônia")
# Configurações da tela

root.bind("<Escape>", sair_tela_cheia)

largura_tela = root.winfo_screenwidth()
altura_tela = root.winfo_screenheight()


# Configurações adicionais da janela
root.title("Sorteio Sicoob")
root.attributes("-fullscreen", True)  # Modo tela cheia

# Carrega as imagens
imagens_paths = [path_user+"OneDrive - Sicoob/Documentos - Sicoob UniRondônia/27. Atualizações Python/Sorteio Sicoob/total.png",path_user+"OneDrive - Sicoob/Documentos - Sicoob UniRondônia/27. Atualizações Python/Sorteio Sicoob/cell.png", path_user+"OneDrive - Sicoob/Documentos - Sicoob UniRondônia/27. Atualizações Python/Sorteio Sicoob/moto.png","carro.png"] 

imagens = [Image.open(path) for path in imagens_paths]

# Redimensiona as imagens para as dimensões da tela
largura_tela = root.winfo_screenwidth()
altura_tela = root.winfo_screenheight()
imagens = [imagem.resize((largura_tela, altura_tela)) for imagem in imagens]

imagens_tk = [ImageTk.PhotoImage(imagem) for imagem in imagens]

# Cria um Canvas em tela cheia
canvas = tk.Canvas(root, width=largura_tela, height=altura_tela)
canvas.pack(fill=tk.BOTH, expand=True)

# Adiciona a imagem do carrossel ao Canvas
imagem_carrossel = imagens_tk[0]  # Use a primeira imagem como carrossel inicial
carrossel_id = canvas.create_image(0, 0, anchor=tk.NW, image=imagem_carrossel)

# Cria bolinhas de navegação
bolinhas = []
for i in range(len(imagens)):
    bolinha = tk.Label(root, text=f"●", font=("Helvetica", 16), cursor="hand2",bg="#999999")
    bolinha.bind("<Button-1>", lambda event, idx=i: selecionar_imagem(idx))
    
    # Ajuste as coordenadas das bolinhas acima do carrossel
    bolinha.place(x=468 + i * 20, y=907 - 30)
    
    bolinhas.append(bolinha)

# Função para atualizar o carrossel com base no índice da imagem
def atualizar_carrossel(idx):
    imagem_atual = imagens_tk[idx]
    canvas.itemconfig(carrossel_id, image=imagem_atual)

    # Atualiza a cor da bolinha selecionada
    for i, bolinha in enumerate(bolinhas):
        cor = "black" if i == idx else "white"
        bolinha.configure(fg=cor)

# Função para selecionar uma imagem ao clicar em uma bolinha
def selecionar_imagem(idx):
    atualizar_carrossel(idx)

# Carrega automaticamente a primeira imagem ao iniciar a tela
atualizar_carrossel(0)

# Definir a cor em hexadecimal (#093640)
cor_hexadecimal = "#093640"




def alterar_cor_ao_clicar():
    botao.config(bg="#138FAB")

# Função para voltar à cor original ao liberar o botão
def voltar_cor_ao_liberar(event=None):
    botao.config(bg=cor_hexadecimal)

# Variável compartilhada entre threads para indicar o estado do segundo botão
segundo_botao_clicado = tk.BooleanVar()
segundo_botao_clicado.set(False)

# Criar um botão com o fundo, texto, fonte, bordas arredondadas, cor ao clicar e voltar ao liberar
botao = tk.Button(root, text="INICIAR", command=botao1_clicado, bg=cor_hexadecimal, fg="white", font=("Asap", 32), border=5, relief=tk.GROOVE, padx=10, pady=5, activebackground="#138FAB")
botão_sortear = tk.Button(root, text="SORTEAR", command=botao2_clicado, bg=cor_hexadecimal, fg="white", font=("Asap", 32), border=5, relief=tk.GROOVE, padx=10, pady=5, activebackground="#138FAB")
# Vincular a função de voltar à cor original ao evento de liberação do botão
botão_sortear.bind("<ButtonRelease-1>", voltar_cor_ao_liberar)
botao.bind("<ButtonRelease-1>", voltar_cor_ao_liberar)
botao.place(x=1171, y=579)




root.mainloop()
