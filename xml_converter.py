"""
Conversor XML para CSV/Excel
Desenvolvido por: Marcelo Piloni
GitHub: https://github.com/marcelopiloni
"""

import xml.etree.ElementTree as ET
import pandas as pd
import csv
import os
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import threading
from typing import List, Dict, Any

class XMLConverter:
    def __init__(self):
        self.data = []
        self.headers = set()
    
    def xml_to_dict(self, element, parent_path=""):
        """Converte elemento XML em dicionário, lidando com estruturas aninhadas."""
        """Desenvolvido por Marcelo Piloni"""
        result = {}
        
        # Adiciona atributos do elemento
        if element.attrib:
            for attr, value in element.attrib.items():
                key = f"{parent_path}@{attr}" if parent_path else f"@{attr}"
                result[key] = value
        
        # Processa texto do elemento
        if element.text and element.text.strip():
            text_key = parent_path if parent_path else "text"
            result[text_key] = element.text.strip()
        
        # Processa elementos filhos
        child_groups = {}
        for child in element:
            child_path = f"{parent_path}.{child.tag}" if parent_path else child.tag
            
            # Agrupa elementos filhos com mesmo nome
            if child.tag not in child_groups:
                child_groups[child.tag] = []
            child_groups[child.tag].append(child)
        
        # Processa cada grupo de elementos filhos
        for tag, children in child_groups.items():
            child_path = f"{parent_path}.{tag}" if parent_path else tag
            
            if len(children) == 1:
                # Elemento único
                child_dict = self.xml_to_dict(children[0], child_path)
                result.update(child_dict)
            else:
                # Múltiplos elementos com mesmo nome
                for i, child in enumerate(children):
                    indexed_path = f"{child_path}[{i}]"
                    child_dict = self.xml_to_dict(child, indexed_path)
                    result.update(child_dict)
        
        return result
    
    def parse_xml_file(self, xml_file_path: str, root_element: str = None, progress_callback=None):
        """Carrega e processa arquivo XML."""
        self.data = []
        self.headers = set()
        
        if progress_callback:
            progress_callback("Carregando arquivo XML...")
        
        tree = ET.parse(xml_file_path)
        root = tree.getroot()
        
        if progress_callback:
            progress_callback(f"Elemento raiz encontrado: {root.tag}")
        
        # Se não especificado, usa elementos filhos diretos da raiz
        if root_element is None:
            # Tenta encontrar elementos repetidos (registros)
            child_tags = [child.tag for child in root]
            if child_tags:
                # Pega o primeiro tipo de elemento filho como registro
                root_element = child_tags[0]
            else:
                # Se não há filhos, usa a própria raiz
                root_element = root.tag
        
        # Coleta todos os elementos do tipo especificado
        if root_element == root.tag:
            elements = [root]
        else:
            elements = root.findall(f".//{root_element}")
            if not elements:
                elements = [child for child in root if child.tag == root_element]
        
        if progress_callback:
            progress_callback(f"Encontrados {len(elements)} registros")
        
        # Converte cada elemento para dicionário
        for i, element in enumerate(elements):
            if progress_callback and i % 100 == 0:
                progress_callback(f"Processando registro {i+1}/{len(elements)}")
            
            record = self.xml_to_dict(element)
            if record:
                self.data.append(record)
                self.headers.update(record.keys())
        
        return len(self.data), len(self.headers)
    
    def to_csv(self, output_file: str, delimiter: str = ','):
        """Exporta dados para CSV."""
        if not self.data:
            raise Exception("Nenhum dado para exportar!")
        
        sorted_headers = sorted(self.headers)
        
        with open(output_file, 'w', newline='', encoding='utf-8') as csvfile:
            writer = csv.DictWriter(csvfile, fieldnames=sorted_headers, delimiter=delimiter)
            writer.writeheader()
            
            for row in self.data:
                complete_row = {header: row.get(header, '') for header in sorted_headers}
                writer.writerow(complete_row)
    
    def to_excel(self, output_file: str):
        """Exporta dados para Excel (.xlsx)."""
        if not self.data:
            raise Exception("Nenhum dado para exportar!")
        
        df = pd.DataFrame(self.data)
        df = df.reindex(sorted(df.columns), axis=1)
        df.to_excel(output_file, index=False, engine='openpyxl')

class XMLConverterGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Conversor XML para CSV/Excel - Marcelo Piloni")
        self.root.geometry("600x520")
        self.root.configure(bg='#f0f0f0')
        
        # Variáveis
        self.xml_file = tk.StringVar()
        self.output_path = tk.StringVar()
        self.csv_var = tk.BooleanVar(value=True)
        self.excel_var = tk.BooleanVar(value=True)
        self.delimiter_var = tk.StringVar(value=',')
        
        self.converter = XMLConverter()
        
        self.create_widgets()
        
        # Centraliza a janela
        self.center_window()
    
    def center_window(self):
        """Centraliza a janela na tela."""
        self.root.update_idletasks()
        x = (self.root.winfo_screenwidth() // 2) - (self.root.winfo_width() // 2)
        y = (self.root.winfo_screenheight() // 2) - (self.root.winfo_height() // 2)
        self.root.geometry(f"+{x}+{y}")
    
    def create_widgets(self):
        """Cria os elementos da interface."""
        # Título
        title_frame = tk.Frame(self.root, bg='#f0f0f0')
        title_frame.pack(pady=20)
        
        title_label = tk.Label(title_frame, text="Conversor XML para CSV/Excel", 
                              font=('Arial', 16, 'bold'), bg='#f0f0f0', fg='#333')
        title_label.pack()
        
        subtitle_label = tk.Label(title_frame, text="Selecione um arquivo XML para converter", 
                                 font=('Arial', 10), bg='#f0f0f0', fg='#666')
        subtitle_label.pack()
        
        # Frame principal
        main_frame = tk.Frame(self.root, bg='#f0f0f0')
        main_frame.pack(padx=20, pady=10, fill='both', expand=True)
        
        # Seção de arquivo XML
        xml_frame = tk.LabelFrame(main_frame, text="Arquivo XML", font=('Arial', 10, 'bold'), 
                                 bg='#f0f0f0', fg='#333')
        xml_frame.pack(fill='x', pady=10)
        
        xml_path_frame = tk.Frame(xml_frame, bg='#f0f0f0')
        xml_path_frame.pack(fill='x', padx=10, pady=10)
        
        self.xml_entry = tk.Entry(xml_path_frame, textvariable=self.xml_file, font=('Arial', 9))
        self.xml_entry.pack(side='left', fill='x', expand=True)
        
        xml_button = tk.Button(xml_path_frame, text="Procurar...", command=self.browse_xml_file,
                              bg='#4CAF50', fg='white', font=('Arial', 9))
        xml_button.pack(side='right', padx=(10, 0))
        
        # Seção de formato de saída
        format_frame = tk.LabelFrame(main_frame, text="Formatos de Saída", font=('Arial', 10, 'bold'),
                                    bg='#f0f0f0', fg='#333')
        format_frame.pack(fill='x', pady=10)
        
        format_inner = tk.Frame(format_frame, bg='#f0f0f0')
        format_inner.pack(padx=10, pady=10)
        
        csv_check = tk.Checkbutton(format_inner, text="CSV", variable=self.csv_var,
                                  font=('Arial', 10), bg='#f0f0f0')
        csv_check.pack(side='left', padx=(0, 20))
        
        excel_check = tk.Checkbutton(format_inner, text="Excel (.xlsx)", variable=self.excel_var,
                                    font=('Arial', 10), bg='#f0f0f0')
        excel_check.pack(side='left')
        
        # Opções CSV
        csv_options_frame = tk.Frame(format_frame, bg='#f0f0f0')
        csv_options_frame.pack(fill='x', padx=10, pady=(0, 10))
        
        tk.Label(csv_options_frame, text="Delimitador CSV:", font=('Arial', 9), 
                bg='#f0f0f0').pack(side='left')
        
        delimiter_combo = ttk.Combobox(csv_options_frame, textvariable=self.delimiter_var,
                                      values=[',', ';', '\t', '|'], width=5, font=('Arial', 9))
        delimiter_combo.pack(side='left', padx=(5, 0))
        
        # Seção de saída
        output_frame = tk.LabelFrame(main_frame, text="Local de Saída", font=('Arial', 10, 'bold'),
                                    bg='#f0f0f0', fg='#333')
        output_frame.pack(fill='x', pady=10)
        
        output_path_frame = tk.Frame(output_frame, bg='#f0f0f0')
        output_path_frame.pack(fill='x', padx=10, pady=10)
        
        self.output_entry = tk.Entry(output_path_frame, textvariable=self.output_path, font=('Arial', 9))
        self.output_entry.pack(side='left', fill='x', expand=True)
        
        output_button = tk.Button(output_path_frame, text="Escolher...", command=self.browse_output_folder,
                                 bg='#2196F3', fg='white', font=('Arial', 9))
        output_button.pack(side='right', padx=(10, 0))
        
        # Botão de conversão
        button_frame = tk.Frame(main_frame, bg='#f0f0f0')
        button_frame.pack(pady=20)
        
        self.convert_button = tk.Button(button_frame, text="CONVERTER ARQUIVO", 
                                       command=self.start_conversion,
                                       bg='#FF5722', fg='white', font=('Arial', 12, 'bold'),
                                       padx=30, pady=10)
        self.convert_button.pack()
        
        # Barra de progresso
        self.progress = ttk.Progressbar(main_frame, mode='indeterminate')
        self.progress.pack(fill='x', pady=10)
        
        # Área de log
        log_frame = tk.LabelFrame(main_frame, text="Status", font=('Arial', 10, 'bold'),
                                 bg='#f0f0f0', fg='#333')
        log_frame.pack(fill='both', expand=True, pady=10)
        
        self.log_text = tk.Text(log_frame, height=8, font=('Consolas', 9), bg='white')
        log_scrollbar = tk.Scrollbar(log_frame, orient='vertical', command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=log_scrollbar.set)
        
        self.log_text.pack(side='left', fill='both', expand=True, padx=(10, 0), pady=10)
        log_scrollbar.pack(side='right', fill='y', padx=(0, 10), pady=10)
        
        # Rodapé com créditos
        credits_frame = tk.Frame(main_frame, bg='#f0f0f0')
        credits_frame.pack(side='bottom', fill='x', pady=(10, 0))
        
        credits_label = tk.Label(credits_frame, 
                               text="Desenvolvido por Marcelo Piloni | GitHub: https://github.com/marcelopiloni",
                               font=('Arial', 8), bg='#f0f0f0', fg='#888')
        credits_label.pack()
        
        # Log inicial
        self.log("Conversor iniciado. Selecione um arquivo XML para começar.")
        self.log("Desenvolvido por Marcelo Piloni - GitHub: https://github.com/marcelopiloni")
    
    def log(self, message):
        """Adiciona mensagem ao log."""
        self.log_text.insert(tk.END, f"{message}\n")
        self.log_text.see(tk.END)
        self.root.update()
    
    def browse_xml_file(self):
        """Abre diálogo para seleção do arquivo XML."""
        filename = filedialog.askopenfilename(
            title="Selecionar arquivo XML",
            filetypes=[("Arquivos XML", "*.xml"), ("Todos os arquivos", "*.*")]
        )
        if filename:
            self.xml_file.set(filename)
            # Define pasta de saída automaticamente
            if not self.output_path.get():
                self.output_path.set(os.path.dirname(filename))
            self.log(f"Arquivo selecionado: {os.path.basename(filename)}")
    
    def browse_output_folder(self):
        """Abre diálogo para seleção da pasta de saída."""
        folder = filedialog.askdirectory(title="Escolher pasta de saída")
        if folder:
            self.output_path.set(folder)
            self.log(f"Pasta de saída: {folder}")
    
    def start_conversion(self):
        """Inicia o processo de conversão em thread separada."""
        if not self.xml_file.get():
            messagebox.showerror("Erro", "Selecione um arquivo XML!")
            return
        
        if not self.csv_var.get() and not self.excel_var.get():
            messagebox.showerror("Erro", "Selecione pelo menos um formato de saída!")
            return
        
        if not self.output_path.get():
            messagebox.showerror("Erro", "Escolha uma pasta de saída!")
            return
        
        # Desabilita botão e inicia progresso
        self.convert_button.config(state='disabled')
        self.progress.start()
        
        # Executa conversão em thread separada
        thread = threading.Thread(target=self.convert_file)
        thread.daemon = True
        thread.start()
    
    def convert_file(self):
        """Executa a conversão do arquivo."""
        try:
            xml_path = self.xml_file.get()
            output_dir = self.output_path.get()
            base_name = os.path.splitext(os.path.basename(xml_path))[0]
            
            self.log("Iniciando conversão...")
            
            # Processa arquivo XML
            num_records, num_fields = self.converter.parse_xml_file(
                xml_path, progress_callback=self.log
            )
            
            self.log(f"✓ Processamento concluído: {num_records} registros, {num_fields} campos")
            
            # Gera arquivos de saída
            files_created = []
            
            if self.csv_var.get():
                csv_path = os.path.join(output_dir, f"{base_name}.csv")
                self.log("Gerando arquivo CSV...")
                self.converter.to_csv(csv_path, self.delimiter_var.get())
                files_created.append(csv_path)
                self.log(f"✓ CSV criado: {os.path.basename(csv_path)}")
            
            if self.excel_var.get():
                excel_path = os.path.join(output_dir, f"{base_name}.xlsx")
                self.log("Gerando arquivo Excel...")
                self.converter.to_excel(excel_path)
                files_created.append(excel_path)
                self.log(f"✓ Excel criado: {os.path.basename(excel_path)}")
            
            self.log(f"\n🎉 CONVERSÃO CONCLUÍDA COM SUCESSO!")
            self.log(f"Arquivos criados em: {output_dir}")
            
            # Pergunta se quer abrir a pasta
            self.root.after(0, lambda: self.conversion_complete(output_dir, files_created))
            
        except Exception as e:
            error_msg = f"❌ Erro durante conversão: {str(e)}"
            self.log(error_msg)
            self.root.after(0, lambda: messagebox.showerror("Erro", str(e)))
        
        finally:
            # Reabilita interface
            self.root.after(0, self.reset_interface)
    
    def conversion_complete(self, output_dir, files_created):
        """Chamado quando conversão é concluída."""
        result = messagebox.askyesno("Conversão Concluída", 
                                   f"Conversão realizada com sucesso!\n\n"
                                   f"Arquivos criados: {len(files_created)}\n"
                                   f"Local: {output_dir}\n\n"
                                   f"Deseja abrir a pasta de destino?")
        if result:
            try:
                os.startfile(output_dir)  # Windows
            except:
                try:
                    os.system(f'open "{output_dir}"')  # macOS
                except:
                    os.system(f'xdg-open "{output_dir}"')  # Linux
    
    def reset_interface(self):
        """Redefine interface após conversão."""
        self.convert_button.config(state='normal')
        self.progress.stop()

def main():
    root = tk.Tk()
    app = XMLConverterGUI(root)
    root.mainloop()

if __name__ == "__main__":
    main()