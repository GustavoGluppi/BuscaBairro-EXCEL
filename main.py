# Importing libraries
import customtkinter as ctk
from customtkinter import filedialog
from CTkMessagebox import CTkMessagebox
import pandas as pd
import urllib.parse
import requests

# Theme settings
ctk.set_appearance_mode("system")

# Initializing CustomTkinter
app = ctk.CTk()
app.title("BuscaBairro")
app.iconbitmap("icon.ico")

# Size options
app.geometry("800x600")

# Creating default labels
ctk.CTkLabel(app, text="BEM VINDO(A)!", font=("tuple", 18, "underline")).place(x=400, y=20, anchor=ctk.CENTER)
ctk.CTkLabel(app, text="Por favor, selecione no campo abaixo o arquivo que deseja extrair.", font=("tuple", 15)).place(x=400, y=60, anchor=ctk.CENTER)

# Open sheet button
fileTypes = [("Arquivos Excel", "*.xlsx; *.xls")]

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

    # District field
    ctk.CTkLabel(app, text="Selecione a coluna de BAIRRO:").place(x=375, y=240, anchor=ctk.E)
    districtBox = ctk.CTkComboBox(app, values=headers)
    districtBox.set(list(s for s in headers if s.lower() in ["bairro"]) or "Selecione...")
    districtBox.place(x=425, y=240, anchor=ctk.W)

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
    def searchDistricts():
        # Gettin' dropdown values (columns names)
        addressHeader = addressBox.get()
        districtHeader = districtBox.get()
        cityHeader = cityBox.get()
        ufHeader = ufBox.get()

        # If not all fields were selected yet
        if "Selecione..." in [addressHeader, districtHeader, cityHeader, ufHeader]:
            return CTkMessagebox(title="Erro!", message="Por favor, selecione as colunas corretamente.", icon="cancel")
        
        numRows = len(df) # Gettin' num of rows
        
        # Iteraing over each row in table
        for i in range(numRows):
            try:
                city = df[cityHeader][i]
                uf = df[ufHeader][i]
                
                fullAddress = df[addressHeader][i].split(",")
                address = fullAddress[0]
                num = int(fullAddress[1]) or 0

                # Requesting VIACEP for address details
                viaCepUrl = f'https://viacep.com.br/ws/{uf}/{city}/{urllib.parse.quote(address)}/json'
                response = requests.get(viaCepUrl).json()

                # Searching for the correct address by using it number
                district = ""

                if len(response) == 1:
                    district = response[0]["bairro"]
                else:
                    for ad in response:
                        if not ad["complemento"]:
                            continue
                        
                        firstNum = ad["complemento"].split("/")[0]
                        if "até" in firstNum:
                            firstNum = 0
                        else:
                            firstNum = "".join(c for c in firstNum if c.isdigit())

                        lastNum = ad["complemento"].split("/")[-1]
                        if "fim" in lastNum:
                            lastNum = 9999
                        else:
                            lastNum = "".join(c for c in lastNum if c.isdigit())
                    
                        if int(firstNum) <= num and int(lastNum) >= num:
                            district = ad["bairro"]
                            return
                        
                print(f'{viaCepUrl}: {district}')
            except requests.exceptions.JSONDecodeError as error:
                print(f'ERROR: {error}')
            
            finally:
                df.loc[i, districtHeader] = district
                continue
        
        # Asking for save the file
        try:
            with filedialog.asksaveasfile(mode='w', defaultextension=".xlsx", filetypes=[("Arquivos Excel", ".xlsx"), ("Todos os arquivos", "*")]) as file:
                df.to_excel(file.name)
        except AttributeError:
            print("The user cancelled save")


    ctk.CTkButton(app, text="Extrair bairros", fg_color="#30a321", hover_color="#1f7314", command=searchDistricts).place(x=400, y=420, anchor=ctk.CENTER)

            

# Starting application
app.mainloop()