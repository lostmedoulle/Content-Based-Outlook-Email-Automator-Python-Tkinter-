from customtkinter import *
import tkinter as tk
from tkinter import filedialog, messagebox
from PIL import Image
import threading
import os
import pandas as pd
from CTkTable import CTkTable
from core_app_master import OutlookProcessor
import subprocess

class OutlookGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Processeur d'Emails Outlook")
        self.root.geometry("1000x1000")  # Increased size for better layout

        self.processor = None
        self.stop_event = threading.Event()
        self.path_var = tk.StringVar()

        # Load image for buttons
        self.button_image = CTkImage(dark_image=Image.open("company_logo_client.png"), light_image=Image.open("company_logo_client.png"))

        self.create_main_layout()
        self.load_saved_path()

    def create_main_layout(self):
        # Créer et configurer le premier cadre avec une vue par onglets pour la configuration et le tutoriel
        frame_1 = CTkFrame(master=self.root, fg_color="#CD8C67")
        frame_1.grid(row=0, column=0, rowspan=3, sticky="nsew", padx=20, pady=20)

        config_tabview = CTkTabview(master=frame_1)
        config_tabview.pack(expand=True, fill='both', padx=10, pady=10)

        config_tabview.add("Configuration")
        config_tabview.add("Tutoriel")

        # Contenu de l'onglet Configuration
        config_tab = config_tabview.tab("Configuration")
        CTkLabel(master=config_tab, text="Configuration", font=("Arial Bold", 20), justify="left").pack(expand=True, pady=[10, 50])
        
        self.path_entry = CTkEntry(master=config_tab, textvariable=self.path_var, width=400, fg_color='white', text_color='black')
        self.path_entry.pack(pady=10)

        CTkButton(master=config_tab, text="Parcourir", command=self.browse_file, image=self.button_image, corner_radius=32, fg_color="#4158D0", hover_color="#C850C0", border_color="#FFCC70", border_width=2).pack(expand=True, pady=20)
        CTkButton(master=config_tab, text="Charger la configuration", command=self.load_configuration, image=self.button_image, corner_radius=32, fg_color="#4158D0", hover_color="#C850C0", border_color="#FFCC70", border_width=2).pack(expand=True, pady=20)

        # Contenu de l'onglet Tutoriel
        tutorial_tab = config_tabview.tab("Tutoriel")
        CTkLabel(master=tutorial_tab, text="Comment utiliser le processeur d'emails Outlook", font=("Arial", 14), justify="left").pack(expand=True, pady=[10, 50])
        tutorial_text = """
        1. Parcourir pour sélectionner le fichier de configuration Excel.
        2. Charger la configuration.
        3. Fermer Outlook / éviter des conflits de session 
        4. Cliquez sur 'Exécuter une fois' pour traiter les emails une fois.
        5. Cliquez sur 'Exécuter toutes les 10 minutes' pour démarrer le traitement périodique.
        6. Cliquez sur 'Arrêter' pour arrêter le traitement périodique.
        7. Cliquez sur 'Quitter' pour fermer l'application.

        Le journal et le contenu Excel seront affichés dans l'onglet principal.
        """
        CTkLabel(master=tutorial_tab, text=tutorial_text, justify='left').pack(expand=True, padx=20, pady=20)

        # Créer et configurer le deuxième cadre avec une vue par onglets pour le traitement manuel et l'automatisation
        frame_2 = CTkFrame(master=self.root, fg_color="#606190")
        frame_2.grid(row=0, column=1, rowspan=3, sticky="nsew", padx=20, pady=20)

        tabview = CTkTabview(master=frame_2)
        tabview.pack(expand=True, fill='both', padx=10, pady=10)

        tabview.add("Traitement manuel")
        tabview.add("Automatisation")

        # Contenu de l'onglet Traitement manuel
        manual_frame = tabview.tab("Traitement manuel")
        CTkLabel(master=manual_frame, text="Traitement manuel", font=("Arial Bold", 20), justify="left").pack(expand=True, pady=(30, 15))
        CTkButton(master=manual_frame, text="Exécuter une fois", command=self.execute_once, image=self.button_image, corner_radius=32, fg_color="#4158D0", hover_color="#C850C0", border_color="#FFCC70", border_width=2).pack(expand=True, padx=20, pady=(20, 50))

        # Contenu de l'onglet Automatisation
        automate_frame = tabview.tab("Automatisation")
        CTkLabel(master=automate_frame, text="Automatisation", font=("Arial Bold", 20), justify="left").pack(expand=True, pady=(30, 15))
        CTkButton(master=automate_frame, text="Exécuter chaque 10min", command=self.execute_periodically, image=self.button_image, corner_radius=32, fg_color="#4158D0", hover_color="#C850C0", border_color="#FFCC70", border_width=2).pack(expand=True, pady=20)
        CTkButton(master=automate_frame, text="Arrêter", command=self.stop_execution, image=self.button_image, corner_radius=32, fg_color="#4158D0", hover_color="#C850C0", border_color="#FFCC70", border_width=2).pack(expand=True, pady=20)

        # Créer et configurer la barre de progression
        self.progress_bar = CTkProgressBar(master=self.root)
        self.progress_bar.grid(row=4, column=0, columnspan=2, pady=20, padx=20, sticky="nsew")
        self.progress_bar.set(0)

        # Créer et configurer le cadre défilable pour les journaux et Excel
        scrollable_frame = CTkScrollableFrame(master=self.root, fg_color="#8D6F3A", border_color="#FFCC70", border_width=2,
                                              orientation="vertical", scrollbar_button_color="#FFCC70")
        scrollable_frame.grid(row=3, column=0, columnspan=2, padx=20, pady=20, sticky="nsew")

        # Créer un cadre pour le champ de texte du journal et les barres de défilement
        log_frame = CTkFrame(master=scrollable_frame)
        log_frame.pack(expand=True, fill='both', padx=10, pady=10)

        # Créer et configurer le champ de texte moderne pour les journaux
        self.log_textbox = CTkTextbox(master=log_frame, corner_radius=10, border_color="#FFCC70", fg_color="white", text_color="black", wrap='none')
        self.log_textbox.pack(side='left', expand=True, fill='both')

        # Créer et configurer la barre de défilement verticale pour le champ de texte du journal
        self.log_scrollbar_y = CTkScrollbar(master=log_frame, command=self.log_textbox.yview, orientation='vertical')
        self.log_scrollbar_y.pack(side='right', fill='y')
        self.log_textbox.configure(yscrollcommand=self.log_scrollbar_y.set)

        # Créer et configurer la barre de défilement horizontale pour le champ de texte du journal
        self.log_scrollbar_x = CTkScrollbar(master=scrollable_frame, command=self.log_textbox.xview, orientation='horizontal')
        self.log_scrollbar_x.pack(side='bottom', fill='x')
        self.log_textbox.configure(xscrollcommand=self.log_scrollbar_x.set)

        # Créer et configurer le cadre pour le contenu Excel
        self.table_frame = CTkFrame(master=scrollable_frame)
        self.table_frame.pack(expand=True, fill='both', padx=10, pady=10)

        # Créer et configurer le bouton Quitter
        CTkButton(master=self.root, text="Quitter", command=self.quit_program, image=self.button_image, corner_radius=32, fg_color="#4158D0", hover_color="#C850C0", border_color="#FFCC70", border_width=2).grid(row=5, column=0, columnspan=3, pady=20, padx=20)

    def browse_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("Fichiers Excel", "*.xlsx")])
        if file_path:
            self.path_var.set(file_path)
            self.save_path(file_path)

    def load_configuration(self):
        file_path = self.path_var.get()
        if file_path:
            self.processor = OutlookProcessor(file_path, log_callback=self.log_message)
            self.processor.load_configuration()
            self.display_excel_content(file_path)
            messagebox.showinfo("Info", "Configuration chargée avec succès.")
        else:
            messagebox.showwarning("Avertissement", "Veuillez sélectionner un fichier Excel valide.")

    def execute_once(self):
        if self.processor:
            self.progress_bar.set(0)
            threading.Thread(target=self.run_with_progress, args=(self.processor.execute_once,)).start()
        else:
            messagebox.showwarning("Avertissement", "Veuillez d'abord charger la configuration.")

    def execute_periodically(self):
        if self.processor:
            self.stop_event.clear()
            self.periodic_thread = threading.Thread(target=self.run_periodically)
            self.periodic_thread.start()
            messagebox.showinfo("Info", "Exécution périodique démarrée.")
        else:
            messagebox.showwarning("Avertissement", "Veuillez d'abord charger la configuration.")

    def run_periodically(self):
        while not self.stop_event.is_set():
            self.run_with_progress(self.processor.execute_once)
            self.stop_event.wait(600)



    def run_with_progress(self, func):
        # Close Outlook before starting the process
        self.close_outlook()

        self.progress_bar.set(0.5)
        func()
        self.progress_bar.set(1.0)
    
        # Reopen Outlook after the process finishes
        self.log_message("Process completed. Outlook was closed and will not be reopened automatically.")

    def close_outlook(self):
        try:
            # Check if Outlook is running
            tasklist_output = subprocess.check_output("tasklist", text=True)
            if "OUTLOOK.EXE" in tasklist_output:
                # Close Outlook if it is running
                subprocess.run(["taskkill", "/F", "/IM", "OUTLOOK.EXE"], check=True)
                self.log_message("Outlook was running and has been closed successfully.")
            else:
                self.log_message("Outlook is not running, so no need to close it.")
        except Exception as e:
            self.log_message(f"Failed to close Outlook: {e}")


    def stop_execution(self):
        if hasattr(self, 'stop_event'):
            self.stop_event.set()
            if hasattr(self, 'periodic_thread'):
                self.periodic_thread.join()
            messagebox.showinfo("Info", "Exécution périodique arrêtée.")
        else:
            messagebox.showwarning("Avertissement", "Aucune exécution périodique à arrêter.")

    def quit_program(self):
        self.stop_execution()
        self.root.quit()

    def log_message(self, message):
        self.log_textbox.configure(state='normal')
        self.log_textbox.insert(tk.END, message + '\n')
        self.log_textbox.configure(state='disabled')
        self.log_textbox.yview(tk.END)

    def save_path(self, path):
        with open("path_save.txt", "w") as f:
            f.write(path)

    def load_saved_path(self):
        if os.path.exists("path_save.txt"):
            with open("path_save.txt", "r") as f:
                saved_path = f.read()
                if os.path.exists(saved_path):
                    self.path_var.set(saved_path)

    def display_excel_content(self, file_path):
        try:
            df = pd.read_excel(file_path)
            self.show_table(df)
        except Exception as e:
            self.log_message(f"Erreur lors de l'affichage du contenu Excel : {e}")
            messagebox.showerror("Erreur", f"Erreur lors de l'affichage du contenu Excel : {e}")

    def show_table(self, df):
        # Supprimer la table précédente si elle existe
        for widget in self.table_frame.winfo_children():
            widget.destroy()

        # Convertir le DataFrame en liste de listes pour la table
        values = df.values.tolist()
        headers = df.columns.tolist()
        values.insert(0, headers)  # Ajouter les en-têtes en tant que première ligne

        table = CTkTable(master=self.table_frame, row=len(values), column=len(headers), values=values)
        table.pack(expand=True, fill="both", padx=20, pady=20)

if __name__ == "__main__":
    root = CTk()
    app = OutlookGUI(root)
    root.mainloop()
