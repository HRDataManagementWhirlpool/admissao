from src.models.sheets import SheetsModel
from src.controllers.sheets import SheetsController
from src.controllers.docusign import DocusignController

import customtkinter
import os
from PIL import Image
import threading

class App(customtkinter.CTk):
    def __init__(self):
        super().__init__()

        self.title("SmartHUB")
        self.geometry("800x400")
        self.minsize(800,400)

        # set grid layout 1x2
        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(1, weight=1)

        # load images with light and dark mode image
        image_path = os.path.join(os.path.dirname(os.path.realpath(__file__)), "media")
        self.logo_image = customtkinter.CTkImage(Image.open(os.path.join(image_path, "CustomTkinter_logo_single.png")), size=(26, 26))
        self.image_icon_image = customtkinter.CTkImage(Image.open(os.path.join(image_path, "image_icon_light.png")), size=(20, 20))
        self.home_image = customtkinter.CTkImage(light_image=Image.open(os.path.join(image_path, "home_dark.png")),
                                                 dark_image=Image.open(os.path.join(image_path, "home_light.png")), size=(20, 20))
        self.add_user_image = customtkinter.CTkImage(light_image=Image.open(os.path.join(image_path, "add_user_dark.png")),
                                                     dark_image=Image.open(os.path.join(image_path, "add_user_light.png")), size=(20, 20))
        self.chat_image = customtkinter.CTkImage(light_image=Image.open(os.path.join(image_path, "chat_dark.png")),
                                                     dark_image=Image.open(os.path.join(image_path, "chat_light.png")), size=(20, 20))

        # create navigation frame
        self.navigation_frame = customtkinter.CTkFrame(self, corner_radius=0)
        self.navigation_frame.grid(row=0, column=0, sticky="nsew")
        self.navigation_frame.grid_rowconfigure(4, weight=1)

        self.navigation_frame_label = customtkinter.CTkLabel(self.navigation_frame, text="  Bem Vindo(a)", image=self.logo_image,
                                                             compound="left", font=customtkinter.CTkFont(size=15, weight="bold"))
        self.navigation_frame_label.grid(row=0, column=0, padx=20, pady=20)

        self.home_button = customtkinter.CTkButton(self.navigation_frame, corner_radius=0, height=40, border_spacing=10, text="Home",
                                                   fg_color="transparent", text_color=("gray10", "gray90"), hover_color=("gray70", "gray30"),
                                                   image=self.home_image, anchor="w", command=self.home_button_event)
        self.home_button.grid(row=1, column=0, sticky="ew")

        self.frame_2_button = customtkinter.CTkButton(self.navigation_frame, corner_radius=0, height=40, border_spacing=10, text="Conferência Admissão",
                                                      fg_color="transparent", text_color=("gray10", "gray90"), hover_color=("gray70", "gray30"),
                                                      image=self.add_user_image, anchor="w", command=self.frame_2_button_event)
        self.frame_2_button.grid(row=2, column=0, sticky="ew")
        
        self.frame_3_button = customtkinter.CTkButton(self.navigation_frame, corner_radius=0, height=40, border_spacing=10, text="Contratos Admissão",
                                                      fg_color="transparent", text_color=("gray10", "gray90"), hover_color=("gray70", "gray30"),
                                                      image=self.chat_image, anchor="w", command=self.frame_3_button_event)
        self.frame_3_button.grid(row=3, column=0, sticky="ew")

        self.appearance_mode_menu = customtkinter.CTkOptionMenu(self.navigation_frame, values=["Light", "Dark", "System"],
                                                                command=self.change_appearance_mode_event)
        self.appearance_mode_menu.grid(row=6, column=0, padx=20, pady=20, sticky="s")

        # create home frame
        self.home_frame = customtkinter.CTkFrame(self, corner_radius=0, fg_color="transparent")
        self.home_frame.grid_columnconfigure(0, weight=1)
        
        self.textbox = customtkinter.CTkTextbox(self.home_frame, width=250)
        self.textbox.grid(row=0, column=0, padx=20, pady=20, sticky="nsew")
        

        # create second frame
        self.second_frame = customtkinter.CTkFrame(self, corner_radius=0, fg_color="transparent")
        self.second_frame.grid_columnconfigure(0, weight=1)
        
        self.textbox2 = customtkinter.CTkTextbox(self.second_frame, width=250)
        self.textbox2.grid(row=0, column=0, padx=20, pady=20, sticky="nsew")
        self.textbox2.insert("0.0", "Como Funciona\n\n" + "Basta iniciar o processo e inserir o diretório onde estão localizados os arquivos de para a realização da conferência\n\nOs arquivos necessários são: relatório de Agência e conta, relatório de conferência, relatório do eSocial, relatório do WorkForce, relatório de dependentes e a cópia do Check List com os REs que deseja verificar.\n\nNo final do processo, será gerado um arquivo excel com os dados conferidos na mesma pasta dos relatórios.")

        self.second_frame_button_1 = customtkinter.CTkButton(self.second_frame, text="Iniciar processo", image=self.add_user_image, compound="right", command=self.start_sheets_process)
        self.second_frame_button_1.grid(row=1, column=0, padx=100, pady=20, sticky="ew")
        
        # create third frame
        self.third_frame = customtkinter.CTkFrame(self, corner_radius=0, fg_color="transparent")
        self.third_frame.grid_columnconfigure(0, weight=1)
        
        self.textbox3 = customtkinter.CTkTextbox(self.third_frame, width=250)
        self.textbox3.grid(row=0, column=0, padx=20, pady=20, sticky="nsew")
        self.textbox3.insert("0.0", "Como Funciona\n\n" + "Basta iniciar o processo e inserir o diretório onde estão localizados os arquivos de para a realização da conferência\n\nOs arquivos necessários são: relatório de Agência e conta, relatório de conferência, relatório do eSocial, relatório do WorkForce, relatório de dependentes e a cópia do Check List com os REs que deseja verificar.\n\nNo final do processo, será gerado um arquivo excel com os dados conferidos na mesma pasta dos relatórios.")

        
        self.tabview = customtkinter.CTkTabview(self.third_frame, width=250, fg_color="transparent")
        self.tabview.grid(row=1, column=0, padx=20, pady=20, sticky="nsew")
        
        self.tabview.add("Aprendizes")
        self.tabview.add("Horistas")
        self.tabview.add("Mensalistas")
        
        self.tabview.tab("Aprendizes").grid_columnconfigure(0, weight=1)  # configure grid of individual tabs
        self.tabview.tab("Horistas").grid_columnconfigure(0, weight=1)
        self.tabview.tab("Mensalistas").grid_columnconfigure(0, weight=1)

        self.tabview_button = customtkinter.CTkButton(self.tabview.tab("Aprendizes"), text="Processo Aprendiz",
                                                      command=self.start_docusign_process)
        self.tabview_button.grid(row=0, column=0, padx=0, pady=10)
        
        self.tabview_button2 = customtkinter.CTkButton(self.tabview.tab("Horistas"), text="Processo Horista",
                                                      command=None)
        self.tabview_button2.grid(row=0, column=0, padx=0, pady=10)
        
        self.tabview_button3 = customtkinter.CTkButton(self.tabview.tab("Mensalistas"), text="Processo Mensalista",
                                                      command=None)
        self.tabview_button3.grid(row=0, column=0, padx=0, pady=10)

        # select default frame
        self.select_frame_by_name("home")

    def select_frame_by_name(self, name):
        # set button color for selected button
        self.home_button.configure(fg_color=("gray75", "gray25") if name == "home" else "transparent")
        self.frame_2_button.configure(fg_color=("gray75", "gray25") if name == "frame_2" else "transparent")
        self.frame_3_button.configure(fg_color=("gray75", "gray25") if name == "frame_3" else "transparent")

        # show selected frame
        if name == "home":
            self.home_frame.grid(row=0, column=1, sticky="nsew")
        else:
            self.home_frame.grid_forget()
        if name == "frame_2":
            self.second_frame.grid(row=0, column=1, sticky="nsew")
        else:
            self.second_frame.grid_forget()
        if name == "frame_3":
            self.third_frame.grid(row=0, column=1, sticky="nsew")
        else:
            self.third_frame.grid_forget()

    def home_button_event(self):
        self.select_frame_by_name("home")

    def frame_2_button_event(self):
        self.select_frame_by_name("frame_2")
        
    def frame_3_button_event(self):
        self.select_frame_by_name("frame_3")

    def change_appearance_mode_event(self, new_appearance_mode):
        customtkinter.set_appearance_mode(new_appearance_mode)
        
    def start_sheets_process(self):
        self.second_frame_button_1.configure(state="disabled")
        self.loading_bar = customtkinter.CTkProgressBar(self.second_frame)
        self.loading_bar.grid(row=2, column=0, padx=150, pady=(10, 10), sticky="ew")
        self.loading_bar.configure(determinate_speed=5)
        self.loading_bar.start()
        threading.Thread(target=self.sheets_process_in_background).start()
        
    def sheets_process_in_background(self):
        try:
            folder_path = customtkinter.filedialog.askdirectory()
            conferencia = SheetsModel(folder_path, 'Conferência').load_sheet()
            contas = SheetsModel(folder_path, 'Contas').load_sheet() 
            dependentes = SheetsModel(folder_path, 'Dependente').load_sheet()
            eSocial = SheetsModel(folder_path, 'eSocial').load_sheet()
            workForce = SheetsModel(folder_path, 'WorkForce').load_sheet()
            checkList, check = SheetsModel(folder_path, 'Check').clone_sheet('Conferência')
            if not all([conferencia, contas, dependentes, eSocial,workForce, checkList]):
                self.status_indicator = customtkinter.CTkLabel(self.second_frame, text="Algumas planilhas não foram encontradas. Verifique os nomes dos arquivos", text_color="orange")
                self.status_indicator.grid(row=3, column=0, pady=10, sticky="nsew")
                return
            SheetsController(checkList, conferencia, contas, dependentes, eSocial, workForce, folder_path, check)
        except:
            self.status_indicator = customtkinter.CTkLabel(self.second_frame, text="Ocorreu um erro!", text_color="red")
        else:
            self.status_indicator = customtkinter.CTkLabel(self.second_frame, text="Processo concluído!", text_color="green")
        finally:
            self.loading_bar.grid_forget()
            self.second_frame_button_1.configure(state="enabled", require_redraw=True)
            self.status_indicator.grid(row=3, column=0, pady=10, sticky="nsew")
            
    def start_docusign_process(self):
        self.tabview_button.configure(state="disabled")
        self.loading_bar = customtkinter.CTkProgressBar(self.tabview.tab("Aprendizes"))
        self.loading_bar.grid(row=1, column=0, padx=150, pady=(10, 10), sticky="ew")
        self.loading_bar.configure(determinate_speed=5)
        self.loading_bar.start()
        threading.Thread(target=self.docusign_process_in_background).start()
        
    def docusign_process_in_background(self):
        try:
            file = customtkinter.filedialog.askopenfile()
            if not all([file]):
                self.status_indicator = customtkinter.CTkLabel(self.tabview.tab("Aprendizes"), text="Arquivo não informado", text_color="orange")
                self.status_indicator.grid(row=2, column=0, pady=10)
                return print("deu erro no arquivo")
            DocusignController.start_docusign_apprentice_process(file.name)
        except:
            self.status_indicator = customtkinter.CTkLabel(self.third_frame, text="Ocorreu um erro!", text_color="red")
        else:
            self.status_indicator = customtkinter.CTkLabel(self.third_frame, text="Processo concluído!", text_color="green")
        finally:
            self.loading_bar.grid_forget()
            self.tabview_button.configure(state="enabled", require_redraw=True)
            self.status_indicator.grid(row=2, column=0, pady=15, sticky="nsew")
            
if __name__ == "__main__":
    app = App()
    app.mainloop()