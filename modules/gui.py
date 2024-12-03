# modules/gui.py
import os
import customtkinter as ctk
from .modules import create_monthly_schedule

class MonthlyScheduleApp:
    def __init__(self):
        # Fenster-Setup
        self.window = ctk.CTk()
        self.window.title("Monatsplan-Generator")
        self.window.geometry("600x400")
        
        # Theme und Aussehen
        ctk.set_appearance_mode("dark")
        ctk.set_default_color_theme("blue")
        
        # Frame für die Eingabefelder
        self.frame = ctk.CTkFrame(self.window)
        self.frame.pack(pady=20, padx=20, fill="both", expand=True)
        
        # Überschrift
        self.label = ctk.CTkLabel(self.frame, text="Monatsplan-Generator", 
                                font=ctk.CTkFont(size=20, weight="bold"))
        self.label.pack(pady=12, padx=10)
        
        # Dateiname Eingabe
        self.filename_frame = ctk.CTkFrame(self.frame)
        self.filename_frame.pack(fill="x", pady=12, padx=10)
        self.filename_label = ctk.CTkLabel(self.filename_frame, text="Excel-Dateiname:")
        self.filename_label.pack(side="left", padx=5)
        self.filename_entry = ctk.CTkEntry(self.filename_frame, placeholder_text="Name (ohne .xlsx)")
        self.filename_entry.pack(side="left", expand=True, fill="x", padx=5)
        
        # Bundesland Auswahl
        self.bundesland_frame = ctk.CTkFrame(self.frame)
        self.bundesland_frame.pack(fill="x", pady=12, padx=10)
        self.bundesland_label = ctk.CTkLabel(self.bundesland_frame, text="Bundesland:")
        self.bundesland_label.pack(side="left", padx=5)
        
        bundeslaender = {
            'SN': 'Sachsen',
            'BY': 'Bayern',
            'BW': 'Baden-Württemberg',
            'BE': 'Berlin',
            'BB': 'Brandenburg',
            'HB': 'Bremen',
            'HH': 'Hamburg',
            'HE': 'Hessen',
            'MV': 'Mecklenburg-Vorpommern',
            'NI': 'Niedersachsen',
            'NW': 'Nordrhein-Westfalen',
            'RP': 'Rheinland-Pfalz',
            'SL': 'Saarland',
            'ST': 'Sachsen-Anhalt',
            'SH': 'Schleswig-Holstein',
            'TH': 'Thüringen'
        }
        
        self.bundesland_var = ctk.StringVar(value="SN")
        self.bundesland_dropdown = ctk.CTkOptionMenu(
            self.bundesland_frame,
            values=list(bundeslaender.keys()),
            variable=self.bundesland_var
        )
        self.bundesland_dropdown.pack(side="left", padx=5)
        
        # Stundenlohn Eingabe
        self.gehalt_frame = ctk.CTkFrame(self.frame)
        self.gehalt_frame.pack(fill="x", pady=12, padx=10)
        self.gehalt_label = ctk.CTkLabel(self.gehalt_frame, text="Stundenlohn (€):")
        self.gehalt_label.pack(side="left", padx=5)
        self.gehalt_entry = ctk.CTkEntry(self.gehalt_frame)
        self.gehalt_entry.insert(0, "12.50")
        self.gehalt_entry.pack(side="left", expand=True, fill="x", padx=5)
        
        # Feiertagszuschlag Checkbox
        self.zuschlag_var = ctk.BooleanVar()
        self.zuschlag_checkbox = ctk.CTkCheckBox(
            self.frame,
            text="Feiertagszuschlag aktivieren",
            variable=self.zuschlag_var
        )
        self.zuschlag_checkbox.pack(pady=12, padx=10)
        
        # Generate Button
        self.generate_button = ctk.CTkButton(
            self.frame,
            text="Monatsplan erstellen",
            command=self.generate_schedule
        )
        self.generate_button.pack(pady=12, padx=10)
        
        # Status Label
        self.status_label = ctk.CTkLabel(
            self.frame,
            text="",
            wraplength=500  # Für besseren Textumbruch bei langen Nachrichten
        )
        self.status_label.pack(pady=12, padx=10)

    def generate_schedule(self):
        try:
            # Eingabevalidierung
            filename = self.filename_entry.get().strip()
            if not filename:
                self.status_label.configure(
                    text="Bitte geben Sie einen Dateinamen ein!",
                    text_color="red"
                )
                return
            
            # Stundenlohn-Validierung
            try:
                gehalt = float(self.gehalt_entry.get().replace(',', '.'))
                if gehalt <= 0:
                    raise ValueError("Stundenlohn muss größer als 0 sein")
            except ValueError as e:
                self.status_label.configure(
                    text="Bitte geben Sie einen gültigen Stundenlohn ein!",
                    text_color="red"
                )
                return
            
            bundesland = self.bundesland_var.get()
            zuschlag = self.zuschlag_var.get()
            
            # Excel erstellen
            erstellte_datei = create_monthly_schedule(
                dateiname=filename,
                bundesland=bundesland,
                use_holiday_bonus=zuschlag,
                stundenlohn=gehalt
            )
            
            # Erfolgsmeldung mit Dateipfad
            absoluter_pfad = os.path.abspath(erstellte_datei)
            self.status_label.configure(
                text=f"Excel-Datei wurde erfolgreich erstellt!\nSpeicherort: {absoluter_pfad}",
                text_color="green"
            )
            
        except Exception as e:
            self.status_label.configure(
                text=f"Ein Fehler ist aufgetreten: {str(e)}",
                text_color="red"
            )

    def run(self):
        self.window.mainloop()