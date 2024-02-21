import json
import os
import sys
import pandas as pd
import tkinter as tk
from PIL import ImageTk, Image
from tkinter import filedialog, messagebox, ttk, simpledialog

# Global variabel for å holde på de lastede dataene og oppsettene
loaded_data = []
field_setups = {}

# Ordbok for å mappe JSON-feltnavn til nye navn, inkludert en representasjon for 'properties'
field_name_mapping = {
    "periods": "Periode",
    "archiveNo": "Arkivnummer",
    "derivedFrom": "Stammer fra",
    "locationIds": "Lokaliteter",
    "yearOfFinds": "Dato",
    "findCategoryIds": "Funnkategori",
    "subNo": "Unr",
    "museumNo": "Museumsnummer",
    "artefacts": "Gjenstand",
    "materials": "Materiale",
    "siteFindNo": "Funnr",
    "artefactCount": "Antall",
    "artefactVariant": "Variant",
    "largestMeasurement": "Største mål",
    "length": "Lengde",
    "artefactDescription": "Beskrivelse",
    # Felt merket med "IKKE IMPORTER" er utelatt
}

# Ordbok for å mappe underfelt av 'places' til norske navn
places_field_mapping = {
    "countyName": "Fylke",
    "cadastralName": "Gårdsnavn",
    "municipalityName": "Kommune",
    "cadastralNo": "Gårdsnummer",
    "properties.no": "Bruksnummer",
}

# Filnavn for lagrede oppsett
field_setups_filename = 'field_setups.json'

# Last inn eksisterende oppsett fra fil
try:
    with open(field_setups_filename, 'r', encoding='utf-8') as setups_file:
        field_setups = json.load(setups_file)
except FileNotFoundError:
    field_setups = {}

def save_field_setup():
    setup_name = simpledialog.askstring("Lagre Oppsett", "Skriv inn et navn for oppsettet:")
    if setup_name:
        selected_fields = [fields_listbox.get(i) for i in fields_listbox.curselection()]
        field_setups[setup_name] = selected_fields
        with open(field_setups_filename, 'w', encoding='utf-8') as setups_file:
            json.dump(field_setups, setups_file, ensure_ascii=False, indent=4)
        update_setups_list()

# Funksjon for å slette et valgt oppsett
def delete_field_setup():
    selected_setup = setups_listbox.get(setups_listbox.curselection())
    if selected_setup and messagebox.askyesno("Slette Oppsett", f"Er du sikker på at du vil slette oppsettet '{selected_setup}'?"):
        del field_setups[selected_setup]
        with open(field_setups_filename, 'w', encoding='utf-8') as setups_file:
            json.dump(field_setups, setups_file, ensure_ascii=False, indent=4)
        update_setups_list()

# Oppdater listen over oppsett
def update_setups_list():
    setups_list.set(list(field_setups.keys()))
# Oppdater tilstanden til eksportknappen basert på om noen felter er valgt
def update_export_button_state(*args):
    selected_fields = fields_listbox.curselection()
    export_button['state'] = tk.NORMAL if selected_fields else tk.DISABLED

# Funksjon for å bla gjennom og velge JSON-fil
def browse_file():
    filename = filedialog.askopenfilename(filetypes=[("JSON files", "*.json")])
    if filename:
        file_path_entry.delete(0, tk.END)
        file_path_entry.insert(0, filename)
        load_json_fields(filename)

# Funksjon for å lagre det nåværende oppsettet av valgte felter
def save_field_setup():
    setup_name = simpledialog.askstring("Lagre Oppsett", "Skriv inn et navn for oppsettet:")
    if setup_name:
        selected_fields = [fields_listbox.get(i) for i in fields_listbox.curselection()]
        field_setups[setup_name] = selected_fields
        with open(field_setups_filename, 'w', encoding='utf-8') as setups_file:
            json.dump(field_setups, setups_file, ensure_ascii=False, indent=4)
        update_setups_list()

# Funksjon for å laste et valgt oppsett
def load_field_setup(event):
    selected_setup_index = setups_listbox.curselection()
    if selected_setup_index:
        selected_setup = setups_listbox.get(selected_setup_index)
        selected_fields = field_setups[selected_setup]
        fields_listbox.selection_clear(0, tk.END)
        for field in selected_fields:
            if field in fields_listbox.get(0, tk.END):
                index = fields_listbox.get(0, tk.END).index(field)
                fields_listbox.selection_set(index)
        update_export_button_state()

# Oppdater listen over oppsett
def update_setups_list():
    setups_list.set(list(field_setups.keys()))

# Funksjon for å laste inn feltene fra JSON-filen og vise dem i GUI
def load_json_fields(filename):
    global loaded_data
    with open(filename, 'r', encoding='utf-8') as file:
        loaded_data = json.load(file)
    # Anta at alle objekter i listen har samme struktur, så bruk det første objektet for å få feltene
    if isinstance(loaded_data, list) and len(loaded_data) > 0:
        # Oppdater listeboksen med feltene definert i field_name_mapping
        field_list = list(field_name_mapping.values())
        # Legg til underfeltene til 'places' definert i places_field_mapping
        field_list.extend(places_field_mapping.values())
        fields.set(field_list)  # Oppdaterer variabelen med den nye listen

# Funksjon for å eksportere de valgte feltene til Excel
def export_to_excel():
    selected_fields = [fields_listbox.get(i) for i in fields_listbox.curselection()]
    if not selected_fields:
        messagebox.showwarning("Ingen felter valgt", "Vennligst velg minst ett felt for å eksportere.")
        return
    # Filtrer dataene basert på valgte felter og konverter lister til strenger
    filtered_data = []
    for item in loaded_data:
        filtered_item = {}
        for field in selected_fields:
            if field in places_field_mapping.values():
                # Finn det opprinnelige feltnavnet basert på det norske navnet
                original_field_name = next(key for key, value in places_field_mapping.items() if value == field)
                if original_field_name == "properties.no":
                    # Spesialbehandling for 'properties.no' for å hente ut 'no' og lage en kommaseparert streng
                    property_numbers = [str(prop['no']) for prop in item.get('places', [{}])[0].get('properties', [])]
                    filtered_item[field] = ', '.join(property_numbers)
                else:
                    # Håndter felt som er inne i 'places' objektet
                    places_info = item.get('places', [{}])[0]
                    filtered_item[field] = places_info.get(original_field_name, '')
            elif field in field_name_mapping.values():
                # Håndter felt som har en direkte mapping
                original_field_name = next((key for key, value in field_name_mapping.items() if value == field), field)
                value = item.get(original_field_name, '')
                # Hvis verdien er en liste, konverter den til en kommaseparert streng
                if isinstance(value, list):
                    filtered_item[field] = ', '.join(map(str, value))
                else:
                    filtered_item[field] = value
            else:
                # For alle andre felt, hent verdien direkte
                filtered_item[field] = item.get(field, '')
        filtered_data.append(filtered_item)
    # Resten av koden for å eksportere til Excel...
    save_filename = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
    if not save_filename:
        return
    df = pd.DataFrame(filtered_data)
    df.to_excel(save_filename, index=False)
    messagebox.showinfo("Eksport fullført", f"Dataen ble eksportert til {save_filename}.")
    pass

def save_user_preference(show_popup):
    with open('user_preference.json', 'w') as pref_file:
        json.dump({'show_popup': show_popup}, pref_file)

def load_user_preference():
    if os.path.exists('user_preference.json'):
        with open('user_preference.json', 'r') as pref_file:
            return json.load(pref_file).get('show_popup', True)
    else:
        return True

def close_popup():
    save_user_preference(not show_popup_var.get())
    popup.destroy()

# Opprett hovedvinduet
root = tk.Tk()
root.title("UNIMUS Kultur - arkeologi - Json til Excel")

# Finn riktig sti for logofilen når programmet er pakket med PyInstaller
if getattr(sys, 'frozen', False):
    # Hvis programmet kjøres som en kjørbar fil (pakket med PyInstaller)
    logo_path = os.path.join(sys._MEIPASS, 'UiT_Segl_Bok_Bla_30px.png')
else:
    # Hvis programmet kjøres som et vanlig Python-skript
    logo_path = r"C:\pythonProject\UiT_Segl_Bok_Bla_30px.png"

# Last inn logoen og sett den som ikonet for vinduet
logo_image = ImageTk.PhotoImage(Image.open(logo_path))
root.iconphoto(False, logo_image)

# Sjekk brukerpreferanse og vis popup om nødvendig
if load_user_preference():
    popup = tk.Toplevel(root)
    popup.title("Om Programmet")
    popup.geometry("400x250")  # Juster størrelsen etter behov
    popup.transient(root)  # Knytt popup-vinduet til hovedvinduet
    popup.grab_set()  # Gjør popupen modal

    # Finn riktig sti for bildet når programmet er pakket med PyInstaller
    if getattr(sys, 'frozen', False):
        # Hvis programmet kjøres som en kjørbar fil (pakket med PyInstaller)
        image_path = os.path.join(sys._MEIPASS, 'UiT_Logo_Bok_2l_Bla_RGB.png')
    else:
        # Hvis programmet kjøres som et vanlig Python-skript
        image_path = r"C:\pythonProject\UiT_Logo_Bok_2l_Bla_RGB.png"

    # Last inn og vis bildet
    image = ImageTk.PhotoImage(Image.open(image_path))
    image_label = tk.Label(popup, image=image)
    image_label.pack(pady=(10, 0))

    tk.Label(popup, text="Dette programmet ble laget av Erik Kjellman, Norges arktiske universitet.\nKontakt: erik.kjellman@uit.no\n\n\nVersjon 0.2\n14.02. 2024", justify=tk.LEFT).pack(pady=(10, 0), padx=10)
    # Legg til mer tekst etter behov

    show_popup_var = tk.BooleanVar(value=False)
    checkbutton = tk.Checkbutton(popup, text="Ikke vis denne meldingen igjen", variable=show_popup_var)
    checkbutton.pack(pady=(10, 0))

    close_button = ttk.Button(popup, text="Lukk", command=close_popup)
    close_button.pack(pady=(10, 0))

    # Sørg for at bildet ikke blir garbage collected
    image_label.image = image

    # Tving fokus til popup-vinduet slik at det vises foran hovedvinduet
    popup.focus_force()

# Opprett en ramme for filvalg
file_frame = ttk.Frame(root)
file_frame.pack(padx=10, pady=10, fill='x', expand=True)

# Opprett en inngangsboks for å vise valgt filsti
file_path_entry = ttk.Entry(file_frame, width=50)
file_path_entry.pack(side=tk.LEFT, expand=True, fill='x', padx=(0, 10))

# Opprett en ramme for oppsett og feltnavn
setup_field_frame = ttk.Frame(root)
setup_field_frame.pack(padx=10, pady=10, fill='both', expand=True)

# Opprett en knapp for å bla gjennom filer
browse_button = ttk.Button(file_frame, text="Bla gjennom", command=browse_file)
browse_button.pack(side=tk.RIGHT)

# Opprett en variabel og listeboks for å vise JSON-felter
fields = tk.Variable(value=[])
fields_listbox = tk.Listbox(setup_field_frame, listvariable=fields, selectmode='multiple', width=50, height=15)
fields_listbox.grid(row=0, column=1, padx=(10, 0), pady=(0, 10))
fields_listbox.bind('<<ListboxSelect>>', update_export_button_state)

# Kall load_json_fields med en eksempel JSON-fil for å initialisere feltlisten
# Erstatt 'example.json' med stien til en faktisk JSON-fil du vil bruke
#load_json_fields('example.json')

# Opprett en ramme for knapper
button_frame = ttk.Frame(root)
button_frame.pack(padx=10, pady=(0, 10), fill='x', expand=True)

# Opprett knapper for å lagre og slette oppsett
save_setup_button = ttk.Button(button_frame, text="Lagre Oppsett", command=save_field_setup)
save_setup_button.pack(side=tk.LEFT, padx=(0, 5))

delete_setup_button = ttk.Button(button_frame, text="Slett Oppsett", command=delete_field_setup)
delete_setup_button.pack(side=tk.LEFT, padx=(5, 0))

# Opprett en variabel og listeboks for å vise lagrede oppsett
setups_list = tk.Variable(value=list(field_setups.keys()))
setups_listbox = tk.Listbox(setup_field_frame, listvariable=setups_list, width=20, height=15)
setups_listbox.grid(row=0, column=0, padx=(0, 10), pady=(0, 10))
setups_listbox.bind('<<ListboxSelect>>', load_field_setup)

# Opprett en knapp for å eksportere til Excel
export_button = ttk.Button(button_frame, text="Eksport til Excel", command=export_to_excel, state=tk.DISABLED)
export_button.pack(side=tk.RIGHT)



# Kjør hovedløkken
root.mainloop()