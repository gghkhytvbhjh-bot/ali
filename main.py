import os
from datetime import datetime
from openpyxl import Workbook, load_workbook

from kivy.config import Config
Config.set('graphics', 'resizable', True)

from kivymd.app import MDApp
from kivymd.uix.screen import MDScreen
from kivymd.uix.boxlayout import MDBoxLayout
from kivymd.uix.toolbar import MDTopAppBar
from kivymd.uix.button import MDRaisedButton, MDFlatButton
from kivymd.uix.textfield import MDTextField
from kivymd.uix.datatables import MDDataTable
from kivymd.uix.dialog import MDDialog
from kivy.metrics import dp


class ControleAidScreen(MDScreen):
    def __init__(self, **kwargs):
        super().__init__(**kwargs)

        if 'ANDROID_ARGUMENT' in os.environ:
            from android.storage import app_storage_path
            self.excel_path = os.path.join(app_storage_path(), "controles_data.xlsx")
        else:
            self.excel_path = "controles_data.xlsx"

        self.init_excel()

        layout = MDBoxLayout(orientation='vertical')

        toolbar = MDTopAppBar(title="ControleAid - Gestion des Controles")
        toolbar.elevation = 4
        layout.add_widget(toolbar)

        form_layout = MDBoxLayout(
            orientation='vertical',
            padding=dp(16),
            spacing=dp(12),
            size_hint_y=None
        )
        form_layout.bind(minimum_height=form_layout.setter('height'))

        self.input_nom = MDTextField(
            hint_text="Nom de l'equipement / Site",
            helper_text="Ex: Extincteur Zone A, Camion 12...",
            helper_text_mode="on_focus"
        )
        self.input_inspecteur = MDTextField(
            hint_text="Nom de l'inspecteur",
            helper_text="Ex: Jean Dupont",
            helper_text_mode="on_focus"
        )
        self.input_remarques = MDTextField(
            hint_text="Remarques / Anomalies constatees",
            multiline=True
        )

        form_layout.add_widget(self.input_nom)
        form_layout.add_widget(self.input_inspecteur)
        form_layout.add_widget(self.input_remarques)

        btn_layout = MDBoxLayout(
            orientation='horizontal',
            spacing=dp(12),
            size_hint_y=None,
            height=dp(50)
        )

        btn_valider = MDRaisedButton(
            text="Enregistrer le Controle",
            on_release=self.sauvegarder_donnees
        )
        btn_rafraichir = MDFlatButton(
            text="Actualiser le Tableau",
            on_release=self.charger_tableau
        )

        btn_layout.add_widget(btn_valider)
        btn_layout.add_widget(btn_rafraichir)
        form_layout.add_widget(btn_layout)

        layout.add_widget(form_layout)

        self.table_container = MDBoxLayout(padding=dp(8))
        layout.add_widget(self.table_container)

        self.add_widget(layout)

        self.data_table = None
        self.charger_tableau()

    def init_excel(self):
        """Cree le fichier Excel s'il n'existe pas encore"""
        if not os.path.exists(self.excel_path):
            wb = Workbook()
            ws = wb.active
            ws.title = "Controles"
            ws.append(["Date", "Equipement", "Inspecteur", "Remarques"])
            wb.save(self.excel_path)

    def sauvegarder_donnees(self, instance):
        """Ajoute une nouvelle ligne dans le fichier Excel"""
        nom = self.input_nom.text.strip()
        inspecteur = self.input_inspecteur.text.strip()
        remarques = self.input_remarques.text.strip()

        if not nom or not inspecteur:
            self.afficher_dialogue("Erreur", "Veuillez remplir au moins le nom et l'inspecteur.")
            return

        nouvelle_ligne = [
            datetime.now().strftime("%Y-%m-%d %H:%M"),
            nom,
            inspecteur,
            remarques if remarques else "R.A.S"
        ]

        try:
            wb = load_workbook(self.excel_path)
            ws = wb.active
            ws.append(nouvelle_ligne)
            wb.save(self.excel_path)

            self.input_nom.text = ""
            self.input_inspecteur.text = ""
            self.input_remarques.text = ""

            self.afficher_dialogue("Succes", "Le controle a bien ete enregistre !")
            self.charger_tableau()

        except Exception as e:
            self.afficher_dialogue("Erreur Systeme", f"Impossible d'ecrire : {str(e)}")

    def charger_tableau(self, *args):
        """Lit le fichier Excel et met a jour le tableau"""
        row_data = []
        try:
            wb = load_workbook(self.excel_path)
            ws = wb.active
            rows = list(ws.iter_rows(min_row=2, values_only=True))
            # Garder les 30 dernières lignes
            derniers = rows[-30:] if len(rows) > 30 else rows
            row_data = [list(map(str, row)) for row in derniers if any(row)]
        except Exception:
            row_data = []

        if self.data_table:
            self.table_container.remove_widget(self.data_table)

        self.data_table = MDDataTable(
            use_pagination=True,
            column_data=[
                ("Date", dp(35)),
                ("Equipement", dp(35)),
                ("Inspecteur", dp(30)),
                ("Remarques", dp(45)),
            ],
            row_data=row_data,
            rows_num=5
        )
        self.table_container.add_widget(self.data_table)

    def afficher_dialogue(self, titre, message):
        """Affiche une boite de dialogue pop-up"""
        dialogue = MDDialog(
            title=titre,
            text=message,
            buttons=[MDFlatButton(text="OK", on_release=lambda x: dialogue.dismiss())]
        )
        dialogue.open()


class ControleAidApp(MDApp):
    def build(self):
        self.theme_cls.primary_palette = "Blue"
        self.theme_cls.theme_style = "Light"
        return ControleAidScreen()


if __name__ == "__main__":
    ControleAidApp().run()
