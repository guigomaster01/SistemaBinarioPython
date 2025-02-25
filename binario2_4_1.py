import os
import re
import pandas as pd
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

# Automação para visualização de ônibus que precisam comunicar com urgência através de WIFI
# 
# Diretório base dos arquivos BIN
directory_base = r"C:\\Mercury.001\\DPT_001\\DATA.UD\\HG\\TG"

# Lista fixa de prefixos
prefixos_fixos = [
    "1406", "1407", "3144", "10008", "10010", "10027", "10031", "10035", "10038", "10154", "10156", "10158", "10160", "10162",
    "10166", "10168", "10170", "10172", "10250", "10254", "10308", "11102", "11103", "11104", "11112", "12001", "12002", "14140",
    "14141", "14187", "14220", "17217", "17218", "17219", "32101", "32102", "32130", "32601", "33102", "33104", "33105", "33109",
    "33110", "33112", "33116", "33117", "33118", "33119", "33120", "33472", "33495", "33514", "81209", "91212", "91213", "133469", "133503"
]

# Expressão regular para extrair data e prefixo dos arquivos
file_pattern = re.compile(r"TG001_(\d{5})_.*_(\d{14})_.*\.BIN")

# Coletando os dados dos arquivos
data_prefixo = []
for root, dirs, files in os.walk(directory_base):
    for file in files:
        match = file_pattern.match(file)
        if match:
            prefixo = match.group(1)
            data_str = match.group(2)[:8]  # Pegando apenas YYYYMMDD
            data_formatada = f"{data_str[:4]}-{data_str[4:6]}-{data_str[6:]}"
            data_prefixo.append([data_formatada, prefixo])

# Criando DataFrame
df = pd.DataFrame(data_prefixo, columns=["Data", "Prefixo"]).drop_duplicates().sort_values(by=["Data", "Prefixo"])

# Botão Sobre
def exibir_sobre():
    messagebox.showinfo("Sobre", "Versão: 2.4.1\nDesenvolvido por: Rodrigo Ferreira Rodrigues\nDPT_001")

# Criando a interface gráfica
root = tk.Tk()
root.title("Consulta de Prefixos - versao: 2.4.1")
root.resizable(False, False)

tk.Label(root, text="Selecione o Mês:").grid(row=0, column=0)
tk.Label(root, text="Selecione o Dia:").grid(row=1, column=0)

meses = sorted(set(df["Data"].str[:7]))
mes_combobox = ttk.Combobox(root, values=meses, state="readonly")
mes_combobox.grid(row=0, column=1)

dia_combobox = ttk.Combobox(root, state="readonly")
dia_combobox.grid(row=1, column=1)

def atualizar_dias(event):
    mes_selecionado = mes_combobox.get()
    dias = sorted(set(df["Data"][df["Data"].str.startswith(mes_selecionado)].tolist()))
    dia_combobox["values"] = dias

tk.Button(root, text="Filtrar", command=lambda: exibir_tabela()).grid(row=2, column=1)
tk.Button(root, text="Exportar para Excel", command=lambda: exportar_excel()).grid(row=4, column=1)
tk.Button(root, text="Sobre", command=exibir_sobre).grid(row=4, column=0)

mes_combobox.bind("<<ComboboxSelected>>", atualizar_dias)

frame_tabela = tk.Frame(root)
frame_tabela.grid(row=3, column=0, columnspan=2, sticky="nsew")

tree = ttk.Treeview(frame_tabela, columns=("Prefixo", "Status"), show="headings", height=20)
tree.heading("Prefixo", text="Prefixo")
tree.heading("Status", text="Status")
scrollbar = ttk.Scrollbar(frame_tabela, orient="vertical", command=tree.yview)
tree.configure(yscrollcommand=scrollbar.set)

tree.pack(side="left", fill="both", expand=True)
scrollbar.pack(side="right", fill="y")

def exibir_tabela():
    for row in tree.get_children():
        tree.delete(row)
    
    dia_selecionado = dia_combobox.get()
    if not dia_selecionado:
        return

    datas_verificadas = [(pd.to_datetime(dia_selecionado) - pd.Timedelta(days=i)).strftime('%Y-%m-%d') for i in range(0, 3)]
    
    for prefixo in prefixos_fixos:
        status = "Faltando"
        tag = "faltando"
        
        if prefixo in df[df["Data"] == datas_verificadas[0]]["Prefixo"].tolist():
            status = "Comunicado"
            tag = "presente"
        elif prefixo in df[df["Data"] == datas_verificadas[1]]["Prefixo"].tolist():
            status = f"Comunicado no dia {datas_verificadas[1]}"
            tag = "comunicado"
        elif prefixo in df[df["Data"] == datas_verificadas[2]]["Prefixo"].tolist():
            status = "Comunicar com urgência"
            tag = "urgente"
        
        tree.insert("", "end", values=(prefixo, status), tags=(tag,))
    
    tree.tag_configure("presente", foreground="green")
    tree.tag_configure("faltando", foreground="red")
    tree.tag_configure("comunicado", foreground="orange", font=("Arial", 10, "bold"))
    tree.tag_configure("urgente", foreground="darkred", font=("Arial", 10, "bold"))

def exportar_excel():
    dia_selecionado = dia_combobox.get()
    if not dia_selecionado:
        messagebox.showwarning("Atenção", "Selecione um dia antes de exportar!")
        return

    datas_verificadas = [(pd.to_datetime(dia_selecionado) - pd.Timedelta(days=i)).strftime('%Y-%m-%d') for i in range(0, 3)]
    
    dados_exportar = []
    for prefixo in prefixos_fixos:
        status = "Faltando"
        if prefixo in df[df["Data"] == datas_verificadas[0]]["Prefixo"].tolist():
            status = "Comunicado"
        elif prefixo in df[df["Data"] == datas_verificadas[1]]["Prefixo"].tolist():
            status = f"Comunicado no dia {datas_verificadas[1]}"
        elif prefixo in df[df["Data"] == datas_verificadas[2]]["Prefixo"].tolist():
            status = "Comunicar com urgência"
        dados_exportar.append([prefixo, status])
    
    df_export = pd.DataFrame(dados_exportar, columns=["Prefixo", "Status"])
    file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel Files", "*.xlsx")], title="Salvar Arquivo")
    if file_path:
        df_export.to_excel(file_path, index=False, engine='openpyxl')
        messagebox.showinfo("Sucesso", f"Arquivo salvo em: {file_path}")


# Botão para exibir informações sobre o programa



root.mainloop()