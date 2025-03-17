# Importing libraries
import customtkinter as ctk
from customtkinter import filedialog
from tkinter import messagebox
import pandas as pd
from geopy.geocoders import Nominatim
geolocator = Nominatim(user_agent="buscaBairro")

# Theme settings
ctk.set_appearance_mode("system")

# Initializing CustomTkinter
app = ctk.CTk()
app.title("BuscaBairro")
app.iconbitmap("icon.ico")

# Size options
appHeight = 600
appWidth = 800

# Center popup
appX = (app.winfo_screenwidth() // 2) - (appWidth // 2)
appY = (app.winfo_screenheight() // 2) - (appHeight // 2)
app.geometry(f'{appWidth}x{appHeight}+{appX}+{appY}')

# Creating default labels
ctk.CTkLabel(app, text="BEM VINDO(A)!", font=("tuple", 18, "underline")).place(x=400, y=20, anchor=ctk.CENTER)
ctk.CTkLabel(app, text="Por favor, selecione no campo abaixo o arquivo que deseja extrair.", font=("tuple", 15)).place(x=400, y=60, anchor=ctk.CENTER)

# Open sheet button
fileTypes = [("Arquivos Excel", "*.xlsx")]

def selectSheet():
    file = filedialog.askopenfilename(filetypes=fileTypes)

    sheet = pd.read_excel(file)
    df = pd.DataFrame(sheet)

    initializeComboBox(df)

    

ctk.CTkButton(app, text="Selecione a planilha...", height=30, command=selectSheet).place(x=400, y=120, anchor=ctk.CENTER)

# Creating the sheet headers dropdowns
def initializeComboBox(df): 
    headers = list(df.keys())

    # Address field
    ctk.CTkLabel(app, text="Selecione a coluna de ENDEREÇO:").place(x=375, y=180, anchor=ctk.E)
    addressBox = ctk.CTkComboBox(app, values=headers)
    addressBox.set(list(s for s in headers if s.lower() in ["endereço", "endereco"]) or "Selecione...")
    addressBox.place(x=425, y=180, anchor=ctk.W)

    # suburb field
    ctk.CTkLabel(app, text="Selecione a coluna de BAIRRO:").place(x=375, y=240, anchor=ctk.E)
    suburbBox = ctk.CTkComboBox(app, values=headers)
    suburbBox.set(list(s for s in headers if s.lower() in ["bairro"]) or "Selecione...")
    suburbBox.place(x=425, y=240, anchor=ctk.W)

    # City field
    ctk.CTkLabel(app, text="Selecione a coluna de MUNICÍPIO:").place(x=375, y=300, anchor=ctk.E)
    cityBox = ctk.CTkComboBox(app, values=headers)
    cityBox.set(list(s for s in headers if s.lower() in ["município", "municipio"]) or "Selecione...")
    cityBox.place(x=425, y=300, anchor=ctk.W)

    # Federative unit field
    ctk.CTkLabel(app, text="Selecione a coluna de UF:").place(x=375, y=360, anchor=ctk.E)
    ufBox = ctk.CTkComboBox(app, values=headers)
    ufBox.set(list(s for s in headers if s.lower() in ["uf"]) or "Selecione...")
    ufBox.place(x=425, y=360, anchor=ctk.W)

    # Extract button
    def searchsuburbs():
        # Gettin' dropdown values (columns names)
        addressHeader = addressBox.get()
        suburbHeader = suburbBox.get()
        cityHeader = cityBox.get()
        ufHeader = ufBox.get()

        # If not all fields were selected yet
        if "Selecione..." in [addressHeader, suburbHeader, cityHeader, ufHeader]:
            return messagebox.showerror("Erro", "Nem todas as colunas foram selecionadas!")
        
        numRows = len(df) # Gettin' num of rows
        
        progressVar = ctk.StringVar()
        progressVar.set(f'Progresso: 0 / {numRows}')
        progressLabel = ctk.CTkLabel(app, textvariable=progressVar)
        progressLabel.place(x=400, y=500, anchor=ctk.CENTER)
        app.update()

        # Iteraing over each row in table
        for i in range(numRows):
            suburb = df[suburbHeader][i] or ""

            city = df[cityHeader][i]
            uf = df[ufHeader][i]
            
            fullAddress = df[addressHeader][i].split(',')
            address = fullAddress[0]
            num = int(fullAddress[1]) or None

            result = geolocator.geocode(f'{address}, {num}, {city}, {uf}', addressdetails=True)

            # Handling errors
            if result and "address" in result.raw:
                suburb = result.raw["address"].get("suburb", "")
            else:
                print("Address not found")

            print(suburb)

            # Updating pandas data frame
            df.loc[i, suburbHeader] = suburb

            progressVar.set(f'Progresso: {i + 1} / {numRows}')
            app.update()

            continue
        
        # Asking for save the file
        try:
            with filedialog.asksaveasfile(mode='w', defaultextension=".xlsx", filetypes=[("Arquivos Excel", ".xlsx"), ("Todos os arquivos", "*")]) as file:
                df.to_excel(file.name)
        except AttributeError:
            print("The user cancelled save")


    ctk.CTkButton(app, text="Extrair bairros", fg_color="#30a321", hover_color="#1f7314", command=searchsuburbs).place(x=400, y=420, anchor=ctk.CENTER)

            

# Starting application
app.mainloop()

def centerWindow(window):
    screen_width = window.winfo_screenwidth()
    screen_height = window.winfo_screenheight()
    x = (screen_width - window.winfo_reqwidth()) // 2
    y = (screen_height - window.winfo_reqheight()) // 2
    window.geometry(f"+{x}+{y}")
