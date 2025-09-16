from tkinter import *
from tkinter import ttk
from tkinter import messagebox
from openpyxl import load_workbook
import os

# ------------------ CONFIGURAÇÃO DA TELA ------------------
tela = Tk()
tela.geometry('950x350')
tela.title('TreeView')

# ------------------ ESTILO ------------------
estilo = ttk.Style()
estilo.theme_use('alt')
estilo.configure('.', font='Arial 14')

# ------------------ TREEVIEW ------------------
tree_view_dados = ttk.Treeview(tela, columns=(1, 2, 3, 4), show='headings')

tree_view_dados.column('1', anchor=CENTER)
tree_view_dados.heading('1', text='ID')

tree_view_dados.column('2', anchor=CENTER)
tree_view_dados.heading('2', text='Nome')

tree_view_dados.column('3', anchor=CENTER)
tree_view_dados.heading('3', text='Idade')

tree_view_dados.column('4', anchor=CENTER)
tree_view_dados.heading('4', text='Sexo')

tree_view_dados.grid(row=3, column=0, columnspan=8, sticky=NSEW)

# ------------------ CAMPOS DE ENTRADA ------------------
Label(text='ID', font='Arial 12').grid(row=1, column=0, sticky='W')
campo_id = Entry(font='Arial 12')
campo_id.grid(row=1, column=1, sticky='W')

Label(text='Nome', font='Arial 12').grid(row=1, column=2, sticky='W')
campo_nome = Entry(font='Arial 12')
campo_nome.grid(row=1, column=3, sticky='W')

Label(text='Idade', font='Arial 12').grid(row=1, column=4, sticky='W')
campo_idade = Entry(font='Arial 12')
campo_idade.grid(row=1, column=5, sticky='W')

Label(text='Sexo', font='Arial 12').grid(row=1, column=6, sticky='W')
campo_sexo = Entry(font='Arial 12')
campo_sexo.grid(row=1, column=7, sticky='W')

# ------------------ LABEL DE LINHAS ------------------
numero_linhas = Label(text='Linhas: ', font='Arial 18')
numero_linhas.grid(row=4, column=0, columnspan=8, sticky='W')


# ------------------ FUNÇÕES ------------------
def contar_linhas(item=''):
    """Conta quantas linhas existem no Treeview"""
    linhas = tree_view_dados.get_children(item)
    numero_linhas.config(text='Linhas: ' + str(len(linhas)))


def add_item():
    """Adiciona item na tabela"""
    if not campo_id.get():
        messagebox.showerror('Erro', 'Digite algo no campo ID')
    elif not campo_nome.get():
        messagebox.showerror('Erro', 'Digite algo no campo Nome')
    elif not campo_idade.get():
        messagebox.showerror('Erro', 'Digite algo no campo Idade')
    elif not campo_sexo.get():
        messagebox.showerror('Erro', 'Digite algo no campo Sexo')
    else:
        tree_view_dados.insert(
            "",
            'end',
            values=(
                campo_id.get(),
                campo_nome.get(),
                campo_idade.get(),
                campo_sexo.get()
            )
        )

        # limpa os campos
        campo_id.delete(0, 'end')
        campo_nome.delete(0, 'end')
        campo_idade.delete(0, 'end')
        campo_sexo.delete(0, 'end')

        contar_linhas()


def deletar_item():
    """Deleta item selecionado da tabela"""
    itens_selecionados = tree_view_dados.selection()
    for item in itens_selecionados:
        tree_view_dados.delete(item)
    contar_linhas()


def alterar_item():
    """Altera item selecionado"""
    if not campo_id.get():
        messagebox.showerror('Erro', 'Digite algo no campo ID')
    elif not campo_nome.get():
        messagebox.showerror('Erro', 'Digite algo no campo Nome')
    elif not campo_idade.get():
        messagebox.showerror('Erro', 'Digite algo no campo Idade')
    elif not campo_sexo.get():
        messagebox.showerror('Erro', 'Digite algo no campo Sexo')
    else:
        item_selecionado = tree_view_dados.selection()[0]
        tree_view_dados.item(
            item_selecionado,
            values=(
                campo_id.get(),
                campo_nome.get(),
                campo_idade.get(),
                campo_sexo.get()
            )
        )

        # limpa os campos
        campo_id.delete(0, 'end')
        campo_nome.delete(0, 'end')
        campo_idade.delete(0, 'end')
        campo_sexo.delete(0, 'end')

def exportar_excel():
    try:
        # A nova lógica para encontrar o caminho correto, tanto no ambiente de desenvolvimento
        # quanto no executável.
        import sys
        
        if getattr(sys, 'frozen', False):
            # Se o programa está rodando como um executável
            base_path = sys._MEIPASS
        else:
            # Se o programa está rodando no ambiente de desenvolvimento
            base_path = os.path.dirname(os.path.abspath(__file__))

        caminho_origem = os.path.join(base_path, 'dados', 'Tratamento_Dados.xlsx')
        
        if not os.path.exists(caminho_origem):
            messagebox.showerror("Erro", f"Arquivo de origem não encontrado: {caminho_origem}")
            return
            
        workbook = load_workbook(filename=caminho_origem)
        sheet = workbook['Vendedores']
    
        for numero_linha in tree_view_dados.get_children():
            linha = tree_view_dados.item(numero_linha)['values']
            sheet.append(linha)
    
        caminho_export = os.path.join(os.path.expanduser("~"), "Documents", "Dados_Exportados.xlsx")
        workbook.save(filename=caminho_export)
        messagebox.showinfo("Sucesso", f"Arquivo salvo em: {caminho_export}")
        
    except Exception as e:
        messagebox.showerror("Erro", f"Ocorreu um erro ao exportar o arquivo: {e}")

# ------------------ BOTÕES ------------------
btn_cadastrar = Button(text='Cadastrar', font='Arial 20', command=add_item)
btn_cadastrar.grid(row=2, column=0, columnspan=2, sticky=NSEW)

btn_deletar = Button(text='Deletar', font='Arial 20', command=deletar_item)
btn_deletar.grid(row=2, column=2, columnspan=2, sticky=NSEW)

btn_alterar = Button(text='Alterar', font='Arial 20', command=alterar_item)
btn_alterar.grid(row=2, column=4, columnspan=2, sticky=NSEW)

btn_exportar = Button(text='Exportar', font='Arial 20', command=exportar_excel)
btn_exportar.grid(row=2, column=6, columnspan=2, sticky=NSEW)

# ------------------ INICIALIZAÇÃO ------------------
contar_linhas()
tela.mainloop()
