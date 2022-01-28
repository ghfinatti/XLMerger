from tkinter import *
import tkinter.filedialog as filedialog
import tkinter as tk
import pandas as pd
from tkinter import messagebox
import os

master = tk.Tk()
master.geometry("500x450")
master.configure(bg='#a0ffc9')
master.tk.call('wm', 'iconphoto', master._w, tk.PhotoImage(file='C:\\excel_merger\\acclogo2.png'))

def choose_output_directory():
    path = filedialog.askdirectory()

    update_text(output_entry, path)

def merge_chosen_files():

    result_filename = filename_entry.get()
    result_filename = result_filename.replace(" ","_")
    output_directory = output_entry.get()
    sheet_input = choose_sheet_entry.get()
    chosenrb = rbvariable.get()

    try:
        sheet_position = int(sheet_input)
        if sheet_position == 0:
            messagebox.showerror("Erro", "Escolha um número de sheet válido")
            return
    except:
        messagebox.showerror("Erro", "Escolha um número de sheet")
        return
    
    if output_directory == "":
        messagebox.showerror("Erro", "Escolha uma pasta para seu arquivo.")
        return "error"
    if result_filename == "":
        messagebox.showerror("Erro", "Escolha um nome para seu arquivo.")
        return "error"
    
    if chosenrb == "sheets":
        file_list = get_files('single')
        messagebox.showinfo("Aviso", "Seus arquivos serão processados, aguarde se o programa parar de responder, você será avisado quando terminar.")
        result_file = merge_sheets(file_list)
    elif chosenrb == "files":
        file_list = get_files('multiple')
        messagebox.showinfo("Aviso", "Seus arquivos serão processados, aguarde se o programa parar de responder, você será avisado quando terminar.")
        result_file = merge_files(file_list)
    elif chosenrb == "choosesheet":
        file_list = get_files('multiple')
        messagebox.showinfo("Aviso", "Seus arquivos serão processados, aguarde se o programa parar de responder, você será avisado quando terminar.")
        result_file = merge_specific_sheet(file_list)
    else:
        messagebox.showerror("Erro",'Escolha uma opção do que fazer com seu(s) arquivo(s)')
        
    result_file_directory = f"{output_directory}/{result_filename}.xlsx"
    result_file['file'].to_excel(result_file_directory, index=False)

    if result_file['rows_exceeded'] == True:        
        messagebox.showinfo("Sucesso", f"Seu arquivo está pronto mas ultrapassou o limite de linhas do excel, juntamos até o arquivo/aba: {result_file['last_read_data']}.")
    else:
        messagebox.showinfo("Sucesso", "Seu arquivo está pronto.")

    os.chdir(output_directory)
    os.system(f'start excel.exe {result_filename}.xlsx')

def get_files(is_single_file):
    file_list = []
    if is_single_file == 'single':
        input_path = filedialog.askopenfilename(filetypes=[("Excel files",".xlsx .xls")])
        file_list.append(input_path)
    elif is_single_file == 'multiple':
        input_path = filedialog.askopenfilename(multiple=True, filetypes=[("Excel files",".xlsx .xls .csv")])
        for i in range(len(input_path)):
            file_list.append(input_path[i])
    return file_list

def merge_sheets(file_list):
    df_list = []
    num_of_rows = 1
    exceeded_xl_rows = False
    try:
        open_file = pd.ExcelFile(file_list[0])
        for i in range (len(open_file.sheet_names)):
            print(f'Juntando a aba{open_file.sheet_names[i]}.')
            df = pd.read_excel(open_file, i, header=None)
            df.insert(0, 'Nome da Aba', open_file.sheet_names[i])
            num_of_rows += len(df.index)
            df.dropna(how='all')
            if num_of_rows > 1048570:
                exceeded_xl_rows = True
                break
            df_list.append(df)
            last_read = open_file.sheet_names[i]
        excel_merged = pd.concat(df_list, ignore_index=True)
    except Exception as e:
        if str(e) != "No objects to concatenate":
            messagebox.showerror("Erro", e)

    return {'file': excel_merged, 'rows_exceeded': exceeded_xl_rows, 'last_read_data': last_read}

def merge_files(file_list):
    num_of_rows = 1
    exceeded_xl_rows = False
    try:
        df_list = []
        for file in file_list:
            file_name = os.path.basename(file)
            print(f'Juntando o arquivo {file_name}.')
            if ".xls" in file_name:    
                df = pd.read_excel(file, header=None)
            elif ".csv" in file_name:
                df = pd.read_csv(file, header=None)
            df.insert(0, 'Nome do Arquivo', file_name)
            num_of_rows += len(df.index)
            df.dropna(how='all')
            if num_of_rows > 1048570:
                exceeded_xl_rows = True
                break
            df_list.append(df)
            last_read = file_name
        excel_merged = pd.concat(df_list, ignore_index=True)
    except Exception as e:
        if str(e) != "No objects to concatenate":
            messagebox.showerror("Erro", e)

    return {'file': excel_merged, 'rows_exceeded': exceeded_xl_rows, 'last_read_data': last_read}

def merge_specific_sheet(file_list):
    df_list = []
    num_of_rows = 1
    exceeded_xl_rows = False
    sheet_position = int(choose_sheet_entry.get())-1
    try:
        for file in file_list:
            file_name = os.path.basename(file)
            print(f'Juntando o arquivo {file_name}.')
            open_file = pd.ExcelFile(file)
            df = pd.read_excel(open_file, sheet_position, header=None)
            df.insert(0, 'Nome do Arquivo', file_name)
            num_of_rows += len(df.index)
            df.dropna(how='all')
            if num_of_rows > 1048570:
                exceeded_xl_rows = True
                break
            df_list.append(df)
            last_read = file_name
        excel_merged = pd.concat(df_list, ignore_index=True)
    except Exception as e:
        if str(e) != "No objects to concatenate":
            messagebox.showerror("Erro", e)

    return {'file': excel_merged, 'rows_exceeded': exceeded_xl_rows, 'last_read_data': last_read}

def update_text(widget, text):
    widget.configure(state="normal")
    widget.delete(0, tk.END)
    widget.insert(0, text)
    widget.configure(state='readonly')

def open_choose_sheet():
    choose_sheet_entry.configure(state="normal")
    choose_sheet_entry.delete(0, tk.END)

def close_choose_sheet():
    choose_sheet_entry.delete(0, tk.END)
    choose_sheet_entry.insert(0, 'Ex: colocar "3" para terceira sheet')
    choose_sheet_entry.configure(state="readonly")
    
master.title("Excel Merger")

rbvariable = StringVar()
rb_sheets = tk.Radiobutton(master, bg="#a0ffc9", fg="#030d5b", font=('Calibri',12,'bold'), text="Unir todas as abas de um arquivo", value="sheets", variable=rbvariable, command=close_choose_sheet)
rb_files = tk.Radiobutton(master, bg="#a0ffc9", fg="#030d5b", font=('Calibri',12,'bold'), text="Unir arquivos (primeira sheet de cada)", value="files", variable=rbvariable, command=close_choose_sheet)
rb_choose_sheet = tk.Radiobutton(master, bg="#a0ffc9", fg="#030d5b", font=('Calibri',12,'bold'), text="Unir arquivos (sheet específica)", value="choosesheet", variable=rbvariable, command=open_choose_sheet)

line = tk.Frame(master, height=1, width=300, bg="#030d5b", relief='groove')
line2 = tk.Frame(master, height=1, width=300, bg="#030d5b", relief='groove')
output_path = tk.Label(master, bg="#a0ffc9", fg="#030d5b", font=('Calibri',12,'bold'), text="Escolha uma pasta de destino para o arquivo:")
output_entry = tk.Entry(master, text="", width=40)
output_entry.configure(state='readonly')
browse2 = tk.Button(master, width=25, bg='white', text="Escolher pasta", command=choose_output_directory)

filename_label = tk.Label(master, bg="#a0ffc9", fg="#030d5b", font=('Calibri',12,'bold'), text="Escolha o nome do novo arquivo:")
filename_entry = tk.Entry(master, text="", width=40)
rb_options_label = tk.Label(master, bg="#a0ffc9", fg="#030d5b", font=('Calibri',12,'bold'), text="Escolha uma opção:")


choose_sheet_label = tk.Label(master, bg="#a0ffc9", fg="#030d5b", font=('Calibri',12,'bold'), text="Escolha a posição da sheet:")
choose_sheet_entry = tk.Entry(master, width=40)
choose_sheet_entry.insert(0, 'Ex: colocar "3" para terceira sheet')
choose_sheet_entry.configure(state='readonly')

begin_button = tk.Button(master, width=25, bg='white', text='Escolher arquivos', command=merge_chosen_files)

output_path.pack(pady=5)
output_entry.pack(pady=5)
browse2.pack(pady=5)

line.pack(pady=10)

filename_label.pack(pady=5)
filename_entry.pack(pady=5)
line2.pack(pady=10)
browse2.pack(pady=5)

rb_options_label.pack(pady=3)
rb_sheets.pack(pady=1)
rb_files.pack(pady=1)
rb_choose_sheet.pack(pady=1)

choose_sheet_label.pack(pady=1)
choose_sheet_entry.pack(pady=5)


begin_button.pack(pady=10)

master.mainloop()
