import tkinter as tk
from barcode import Code128
from barcode.writer import ImageWriter
from docx import Document
from docx.shared import Inches
import os
import shutil

desktop_path = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop')

codigos = []
png_files = []

def gerar_codigo_de_barras():
    # obtém o número digitado pelo usuário
    numero = entrada_numero.get()

    # gera o código de barras com o número fornecido
    codigo = Code128(numero, writer=ImageWriter())

    # salva o código de barras como imagem PNG
    filename = codigo.save(numero)
    file_path = os.path.join(os.getcwd(), filename)

    # adiciona o código de barras à lista
    codigos.append((numero, file_path))
    png_files.append(filename)

    # exibe uma mensagem na tela informando o sucesso da operação
    label_mensagem.config(text=f"O código de barras {numero} foi gerado com sucesso!")

def gerar_documento():
    # cria um novo documento Word
    doc = Document()

    # insere cada código de barras na lista no documento
    for numero, file_path in codigos:
        # adiciona um cabeçalho com o número do código de barras
        doc.add_heading(f"Produto: (INSIRA O NOME)", level=1)

        # adiciona a imagem do código de barras ao documento
        doc.add_picture(file_path, width=Inches(2))

    # salva o documento como um arquivo .docx
    doc.save(os.path.join(desktop_path, "Codigos_de_barras.docx"))

    # exclui os arquivos PNG gerados -------Conferir se ta apagando mesmo apos testar
    for png_file in png_files:
        os.remove(os.path.join(os.getcwd(), png_file))

    # limpa a lista de arquivos PNG gerados
    png_files.clear()


    # exibe uma mensagem na tela informando o sucesso da operação
    label_mensagem.config(text="O documento com os códigos de barras foi gerado com sucesso!")

# cria a janela principal da aplicação
janela = tk.Tk()
janela.title("Gerador de código de barras")

# cria os widgets da janela
label_numero = tk.Label(janela, text="Digite o número para gerar o código de barras:")
entrada_numero = tk.Entry(janela)
botao_gerar = tk.Button(janela, text="Gerar código de barras", command=gerar_codigo_de_barras)
botao_gerar_documento = tk.Button(janela, text="Clique aqui para finalizar e gerar o documento", command=gerar_documento)
label_mensagem = tk.Label(janela, text="")

# posiciona os widgets na janela
label_numero.pack()
entrada_numero.pack()
botao_gerar.pack()
botao_gerar_documento.pack()
label_mensagem.pack()

# inicia o loop principal da aplicação
janela.mainloop()
