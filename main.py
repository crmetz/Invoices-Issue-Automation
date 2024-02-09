import tkinter as tk
from tkinter import simpledialog
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service
from EmissaoNfe import EmissaoNfe
import os
import sys

class ConfirmationPopup:
    def __init__(self, parent, data):
        self.parent = parent
        self.data = data
        self.create_popup()

    def create_popup(self):
        # popup = tk.Toplevel(self.parent.root)
        # popup.title("Confirmação")
        self.popup = tk.Toplevel(self.parent.root)
        self.popup.title("Confirmação")
        
        # Obtém a largura e altura da tela
        screen_width = self.parent.root.winfo_screenwidth()
        screen_height = self.parent.root.winfo_screenheight()

        # Calcula a posição X para o centro da parte direita da tela
        x_position = screen_width - 400  # 400 é a largura da janela de confirmação
        x_position //= 2 # Centraliza na parte direita

        # Calcula a posição Y para o centro da tela
        y_position = screen_height // 2 - 150  # 150 é a metade da altura da janela de confirmação

        # Define a posição da janela de confirmação
        self.popup.geometry(f"400x300+{x_position}+{y_position}")

        self.popup.configure(bg='#EFECEC')

        label = tk.Label(self.popup, text="Confirme a emissão das seguintes Notas Fiscais:", font=('Arial', 12, 'bold'), bg='#EFECEC')
        label.pack(pady=10)

        label = tk.Label(self.popup, text=f"{self.data}", font=('Arial', 12), bg='#EFECEC')
        label.pack(pady=10)

        button_frame = tk.Frame(self.popup, bg='#EFECEC')
        button_frame.pack(pady=10)

        cancel_button = tk.Button(button_frame, text="Cancelar", command=self.on_cancel, fg="#fff", bg='#DD3C4B', font=('Arial', 14, 'bold'))
        cancel_button.grid(row=0, column=1, padx=10)

        ok_button = tk.Button(button_frame, text="OK", command=self.on_ok, bg='#49CA3E', fg="#fff", font=('Arial', 14, 'bold'))
        ok_button.grid(row=0, column=0, padx=10)

    def on_ok(self):
        # Se o usuário clicou em "OK", execute a ação desejada
        self.parent.emitir_action()
        self.parent.driver.quit()
        self.parent.root.destroy()

    def on_cancel(self):
        # Se o usuário clicou em "Cancelar", apenas feche a janela de confirmação
        #self.parent.driver.quit()
        self.popup.destroy()

class EmitirNfeApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Emitir NFEs")
        # Calculate half of the screen width and height
        screen_width = root.winfo_screenwidth()
        screen_height = root.winfo_screenheight()
        half_screen_width = screen_width // 2
        # Set the window size to half of the screen width and height
        self.root.geometry(f"{half_screen_width}x{screen_height}")

        self.root.configure(bg="#EFECEC")

        # Get the screen width
        screen_width = root.winfo_screenwidth()

        # Position Tkinter window on the right side of the screen
        root.geometry(f"+{screen_width//2}+0")

        # Create a frame for Tkinter GUI
        self.gui_frame = tk.Frame(self.root, bg="#EFECEC", padx=20, pady=10)
        self.gui_frame.grid(row=0, column=0, pady=10, padx=20, sticky="nsew")

        # Lista de containers de entrada de página
        self.page_containers = []

        # Chame o método create_widgets para criar os widgets
        self.create_widgets()

        # Defina o ícone da aplicação
        self.set_application_icon()

        self.open_orders_viewer()

        
    def open_orders_viewer(self):
        script_dir = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))

        chrome_driver_path = os.path.join(script_dir, 'chromedriver.exe')
        
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        half_screen_width = screen_width // 2

        service = Service(chrome_driver_path)
        self.driver = webdriver.Chrome(service=service)

        self.driver.set_window_size(half_screen_width, screen_height)
        self.driver.get('https://app.vhsys.com.br/index.php?Secao=Vendas.Pedidos&Modulo=Vendas')

        username_field = self.driver.find_element('id', 'login')
        password_field = self.driver.find_element('id', 'senha')
        username_field.send_keys('silmarmetz@gmail.com')
        password_field.send_keys('Silmar12#')
        password_field.send_keys(Keys.RETURN)

    def create_widgets(self):
        # Configurações de fonte
        font_label = ('Open Sans', 12)
        font_button = ('Open Sans', 14, 'bold')

        # Título
        title_label = tk.Label(self.root, text="Emitir Notas Fiscais", font=font_button, bg="#EFECEC")
        title_label.grid(row=0, column=0, pady=10, padx=20, sticky="w")

        # Container de botões
        buttons_container = tk.Frame(self.root, padx=20, pady=10, bg="#EFECEC")
        buttons_container.grid(row=1, column=0, pady=10, padx=20, sticky="w")

        # Botões (cor azul)
        button_plus = tk.Button(buttons_container, text="+", font=font_button, bg="#FC6736", fg="#fff", command=self.add_page_container)
        button_plus.grid(row=0, column=0, padx=(0, 10))

        button_minus = tk.Button(buttons_container, text="-", font=font_button, bg="#FC6736", fg="#fff", command=self.remove_page_container)
        button_minus.grid(row=0, column=1, padx=(0, 10))

        # button_print = tk.Button(buttons_container, text="Emitir", font=font_button, bg="#FC6736", fg="#fff", command=self.emitir_action)
        # button_print.grid(row=0, column=2, padx=(0, 10))
        button_print = tk.Button(buttons_container, text="Emitir", font=font_button, bg="#FC6736", fg="#fff",
                                 command=self.show_confirmation_popup)
        button_print.grid(row=0, column=2, padx=(0, 10))

        # Adiciona um container inicial
        self.add_page_container()



    def add_page_container(self):
        if len(self.page_containers) < 6:
            font_label = ('Open Sans', 12)
            # Adiciona um novo container de entrada de página
            page_container = tk.Frame(self.root, bg="white", bd=2, relief=tk.GROOVE, borderwidth=2, padx=20, pady=20, border=5, highlightbackground="#FC6736")
            page_container.grid(row=len(self.page_containers) + 2, column=0, pady=10, padx=20, sticky="w")
            
            
            # Adiciona o rótulo com o número da nota
            nota_label = tk.Label(page_container, text=f"Nota {len(self.page_containers) + 1}", font=('Open Sans', 12, 'bold'), bg="white", anchor="w")
            nota_label.grid(row=0, column=0, columnspan=4, pady=5, sticky="w")

            label_page = tk.Label(page_container, text="Página:", font=font_label, bg="white")
            label_page.grid(row=1, column=0, padx=(0, 5), pady=5)

            entry_page = tk.Entry(page_container, font=font_label, width=5)
            entry_page.grid(row=1, column=1, padx=(0, 5), pady=5)

            label_from = tk.Label(page_container, text="Número do Pedido:", font=font_label, bg="white")
            label_from.grid(row=1, column=2, padx=(10, 5), pady=5)

            entry_from = tk.Entry(page_container, font=font_label, width=10)
            entry_from.grid(row=1, column=3, padx=(0, 5), pady=5)

            # Adiciona o novo container à lista
            self.page_containers.append(page_container)
        

    def remove_page_container(self):
        # Remove o último container de entrada de página
        if len(self.page_containers) > 1:
            if self.page_containers:
                container_to_remove = self.page_containers.pop()
                container_to_remove.destroy()

    def set_application_icon(self):
        # Caminho para o arquivo de ícone (.ico)
        icon_path = 'logo.ico'  # Substitua pelo caminho real do seu ícone

        # Define o ícone da aplicação
        self.root.iconbitmap(icon_path)

    def show_confirmation_popup(self):
        # Recupera os valores das páginas e números de pedidos dos containers
        notes_and_orders = [f"Nota {i + 1}: {container.children['!entry2'].get()}" for i, container in enumerate(self.page_containers)]

        # Cria a mensagem a ser exibida no popup
        message = "\n".join(notes_and_orders)

        # Cria e mostra a janela de confirmação
        ConfirmationPopup(self, message)

    def emitir_action(self):
        self.driver.quit()

        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        half_screen_width = screen_width // 2

        # Recupera os valores das páginas e números de pedidos dos containers
        pages = [container.children['!entry'].get() for container in self.page_containers]

        # Lista para armazenar num_inicial e num_final para cada container
        num_pedidos = []


        for container in self.page_containers:
            num_pedido = container.children['!entry2'].get()  # Assume que o segundo Entry é para 'De'
            num_pedidos.append(num_pedido)

        # Cria uma instância de OrdersAutomation e chama run_automation
        emissao_automation = EmissaoNfe(pages, num_pedidos, half_screen_width, screen_height)
        emissao_automation.run_automation()

        self.open_orders_viewer();


if __name__ == "__main__":
    root = tk.Tk()
    app = EmitirNfeApp(root)
    root.mainloop()