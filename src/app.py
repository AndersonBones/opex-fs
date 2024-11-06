from customtkinter import *
from PIL import Image
import threading
from modules.opex import Opex
from modules.formatting import Formatting
from CTkMessagebox import CTkMessagebox
import datetime


class Gui:
    def __init__(self) -> None:

        path = r"F:\BIOMASSA\00. Indicadores Gerenciais\03. Budget Biomassa\2024-2025\OPEX - Gasto Fixo\Budget 24'25"
        self.export_file_name = f"Opex - Detalhado {datetime.date.today().month-1}_{datetime.date.today().year}.xlsx"

        self.output_path_opex = os.path.join(path, self.export_file_name)
    

        self.app = CTk()
        self.app.geometry("380x200")
        self.app.title("Opex SAC")

        self.app.resizable(height=False, width=False)

        self.bg = '#004B93'

        self.title = CTkLabel(
            self.app, 
            text="Opex SAC", 
            font=("Arial", 22), 
            fg_color=self.bg,
            text_color='#FFFFFF',
            compound="left",
            width=380,
            height=40,   
        )

        self.info_directory = CTkLabel(
            self.app,
            text='Selecione a base Opex exportada do SAC',

        )
      

        self.search_label = CTkLabel(
            self.app, 
            text="Selecione a base Opex", 
            font=("Arial", 18), 
            text_color='white',
            compound="center", 
        )

        self.input_path = StringVar()
        self.entry = CTkEntry(
            master=self.app, 
            placeholder_text='selecione o diretório', 
            width=260, 
            height=30,
            textvariable=self.input_path,
        )

        self.search_button = CTkButton(
            self.app,
            text="Procurar",
            width=50,
            corner_radius=5,
            command=self.selectFile,
           
            
        )
        self.search_button.place(x=295, y=90)

        self.run_button = CTkButton(
            self.app,
            text="Processar",
            width=50,
            corner_radius=5, 
            fg_color='#239b56',
            hover_color='#1d8348',
            command=self.run_process

        )    

        self.progress_bar = CTkProgressBar(
            master=self.app,
            orientation='horizontal',
            width=245,
            height=15,
            determinate_speed=10

        )
        self.progress_bar.set(0)
        
        self.title.place(x=0, y=0)

        self.info_directory.place(x=25, y=65)

        self.entry.place(x=25, y=90)

        self.run_button.place(x=25, y=150)

        self.progress_bar.place(x=110, y=156)

    
    def success_popup(self):
        self.msg = CTkMessagebox(
            title='Opex - SAC',
            message="Concluido!", 
            icon="check", 
            option_1="OK",
            width=300,
            height=200,
            sound=True
        )
    

    def error_popup(self, message):
        self.msg = CTkMessagebox(
            title='Opex - SAC',
            message=message, 
            icon="cancel", 
            option_1="Exit",
            width=300,
            height=200,
            sound=True
        )



    def selectFile(self):
        self.username = os.getlogin()
        self.directory_path = filedialog.askopenfilename(
            defaultextension=".xlsx",
            title="Open File",
		    filetypes=(("Excel Files", ".xlsx"),)
        )

        self.input_path.set(self.directory_path)
    

    
    def save_file_popup(self, file_name=''):
        self.output_file = filedialog.asksaveasfilename(
            title='Selecione o diretório',
            defaultextension='.xlsx',
            filetypes=(('Excel Document', '.xlsx'), ('All Files', '*.*')),
            initialfile=file_name
        )

        
        

    def run_opex(self):
        opex = Opex(self.directory_path)
        opex.run()

        self.progress_bar.set(2)

        self.save_file_popup(file_name=self.export_file_name)

        opex.save_file(export_file_path=self.output_file)

        self.success_popup()

    

    def run_process(self):        
        thr1 = threading.Thread(target=self.run_opex, daemon=True)
        thr1.start()


    def run_gui(self):
        self.app.mainloop()
        



if __name__ == '__main__':
    
    Gui().run_gui()
