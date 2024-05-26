# Importa os módulos necessários
import os
import pandas as pd
from tkinter import Tk, filedialog, Button, Label, messagebox

# Define as listas para armazenar imagens renomeadas e não renomeadas como globais
renamed_images = []
not_renamed_images = []

# Função para selecionar o arquivo Excel
def select_excel_file():
    file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
    excel_file_label.config(text=file_path)  # Atualiza o rótulo com o caminho do arquivo selecionado
    return file_path

# Função para selecionar a pasta de imagens
def select_image_folder():
    image_folder = filedialog.askdirectory()
    image_folder_label.config(text=image_folder)  # Atualiza o rótulo com o caminho da pasta selecionada
    return image_folder

# Função para selecionar a pasta de destino
def select_output_folder():
    output_folder = filedialog.askdirectory()
    output_folder_label.config(text=output_folder)  # Atualiza o rótulo com o caminho da pasta selecionada
    return output_folder

# Função principal para renomear as imagens
def rename_images():
    global renamed_images, not_renamed_images  # Adiciona as listas globais aqui

    # Obtém os caminhos dos arquivos e pastas selecionados
    file_path = excel_file_label.cget("text")
    image_folder = image_folder_label.cget("text")
    output_folder = output_folder_label.cget("text")

    # Verifica se foram selecionados todos os itens necessários
    if not file_path:
        set_status_message("Nenhum arquivo selecionado. Por favor, selecione um arquivo Excel.")
        return
    if not image_folder:
        set_status_message("Nenhuma pasta de imagens selecionada. Por favor, selecione uma pasta de imagens.")
        return
    if not output_folder:
        set_status_message("Nenhuma pasta de destino selecionada. Por favor, selecione uma pasta de destino.")
        return

    # Lê o arquivo Excel
    df = pd.read_excel(file_path)

    # Verifica se as colunas necessárias estão presentes no arquivo Excel
    if "NOME ATUAL" not in df.columns or "NOME ALTERADO" not in df.columns:
        set_status_message("As colunas NOME ATUAL e NOME ALTERADO são necessárias no arquivo Excel.")
        return

    # Limpa as listas de imagens renomeadas e não renomeadas
    renamed_images = []
    not_renamed_images = []

    # Itera sobre as linhas do DataFrame do Excel
    for index, row in df.iterrows():
        current_image_name = row["NOME ATUAL"]
        new_image_name = row["NOME ALTERADO"]

        # Verifica se os nomes das imagens são strings
        if isinstance(current_image_name, str) and isinstance(new_image_name, str):
            # Constrói os caminhos das imagens
            current_image_path_no_ext = os.path.join(image_folder, os.path.splitext(current_image_name)[0])
            current_image_path_with_ext = os.path.join(image_folder, current_image_name)

            # Verifica se as imagens existem na pasta de imagens
            if os.path.exists(current_image_path_no_ext):
                current_image_path = current_image_path_no_ext
            elif os.path.exists(current_image_path_with_ext):
                current_image_path = current_image_path_with_ext
            else:
                # Adiciona à lista de imagens não renomeadas se a imagem não for encontrada
                not_renamed_images.append((current_image_name, f"Imagem não encontrada na pasta de imagens: {current_image_name}"))
                continue

            # Constrói o novo caminho da imagem
            new_image_path = os.path.join(output_folder, new_image_name)

            try:
                # Tenta renomear a imagem
                os.rename(current_image_path, new_image_path)
                renamed_images.append((current_image_name, new_image_name))  # Adiciona à lista de imagens renomeadas
            except Exception as e:
                # Adiciona à lista de imagens não renomeadas se ocorrer um erro
                not_renamed_images.append((current_image_name, f"Erro ao renomear a imagem {current_image_name}: {e}"))

    # Escreve um resumo das imagens renomeadas e não renomeadas em um arquivo de texto
    write_summary_to_file(renamed_images, not_renamed_images)

    # Exibe uma mensagem de conclusão para o usuário
    messagebox.showinfo("Conclusão", "Processo concluído. Verifique o arquivo de texto para detalhes.")
    set_status_message("Processo concluído. Verifique o arquivo de texto para detalhes.")

# Função para escrever um resumo das imagens renomeadas e não renomeadas em um arquivo de texto
def write_summary_to_file(renamed_images, not_renamed_images):
    with open("renaming_summary.txt", "w") as file:
        file.write("Imagens Renomeadas:\n")
        for old_name, new_name in renamed_images:
            file.write(f"{old_name} -> {new_name}\n")
        file.write("\n")
        file.write("Imagens Não Renomeadas:\n")
        for image, error_message in not_renamed_images:
            file.write(f"{image}: {error_message}\n")

# Função para definir uma mensagem de status na interface gráfica
def set_status_message(message):
    status_label.config(text=message)

# Função para salvar o resumo das imagens renomeadas e não renomeadas em um arquivo de texto
def save_text_file():
    file_path = filedialog.asksaveasfilename(defaultextension=".txt", filetypes=[("Text files", "*.txt")])
    if file_path:
        with open(file_path, "w") as file:
            file.write("Imagens Renomeadas:\n")
            for old_name, new_name in renamed_images:
                file.write(f"{old_name} -> {new_name}\n")
            file.write("\n")
            file.write("Imagens Não Renomeadas:\n")
            for image, error_message in not_renamed_images:
                file.write(f"{image}: {error_message}\n")

# Configuração da interface gráfica
root = Tk()
root.title("Renomear Imagens")
root.geometry("600x350")

# Botão para selecionar o arquivo Excel
excel_button = Button(root, text="Arquivo Excel", command=select_excel_file, width=20)
excel_button.pack(pady=5)

# Rótulo para exibir o caminho do arquivo Excel selecionado
excel_file_label = Label(root, text="Nenhum arquivo selecionado.")
excel_file_label.pack()

# Botão para selecionar a pasta de imagens
image_button = Button(root, text="Pasta de Imagens", command=select_image_folder, width=20)
image_button.pack(pady=5)

# Rótulo para exibir o caminho da pasta de imagens selecionada
image_folder_label = Label(root, text="Nenhuma pasta selecionada.")
image_folder_label.pack()

# Botão para selecionar a pasta de destino
output_button = Button(root, text="Pasta de Destino", command=select_output_folder, width=20)
output_button.pack(pady=5)

# Rótulo para exibir o caminho da pasta de destino selecionada
output_folder_label = Label(root, text="Nenhuma pasta selecionada.")
output_folder_label.pack()

# Botão para renomear as imagens
rename_button = Button(root, text="Renomear Imagens", command=rename_images, width=20)
rename_button.pack(pady=10)

# Botão para salvar o resumo das imagens renomeadas e não renomeadas em um arquivo de texto
save_button = Button(root, text="Salvar logs", command=save_text_file, width=20)
save_button.pack(pady=10)

# Rótulo para exibir mensagens de status
status_label = Label(root, text="")
status_label.pack()

# Inicia o loop principal da interface gráfica
root.mainloop()
