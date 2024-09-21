import sys
import json
import os
import pandas as pd
from PyQt5.QtWidgets import (QApplication, QMainWindow, QCheckBox, QPushButton, 
                             QGridLayout, QWidget, QMessageBox, QScrollArea, 
                             QVBoxLayout, QGroupBox, QFileDialog, QLabel, QLineEdit)
from PyQt5.QtGui import QPalette, QColor

# Load the terms from the provided JSON files
with open('terms.json', 'r', encoding='utf-8') as file:
    categories = json.load(file)

with open('termsDescription.json', 'r', encoding='utf-8') as file:
    terms_description = json.load(file)

with open('manualMappings.json', 'r', encoding='utf-8') as file:
    manual_mappings = json.load(file)

# Create a French-to-English mapping for terms
terms_fr_to_en = {}
for category, terms in categories.items():
    for en_term, fr_term in terms.items():
        terms_fr_to_en[fr_term.lower()] = en_term  # Map French terms to English

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Entrez une description de vos besoins en son")

        # Set the dark theme (dark grey background and white text)
        palette = QPalette()
        palette.setColor(QPalette.Window, QColor(40, 40, 40))  # Dark grey background
        palette.setColor(QPalette.WindowText, QColor(255, 255, 255))  # White text
        self.setPalette(palette)

        # Create central widget and layout
        self.central_widget = QWidget()
        self.setCentralWidget(self.central_widget)
        self.main_layout = QVBoxLayout(self.central_widget)

        # Description input field
        self.description_label = QLabel("Description :")
        self.description_label.setStyleSheet("color: white;")
        self.main_layout.addWidget(self.description_label)

        self.description_input = QLineEdit()
        self.description_input.setStyleSheet("background-color: #333; color: white;")
        self.main_layout.addWidget(self.description_input)

        # Validation button for description input
        self.validate_button = QPushButton("Valider la description")
        self.validate_button.setStyleSheet("background-color: #333; color: white;")
        self.validate_button.clicked.connect(self.on_validate_description)
        self.main_layout.addWidget(self.validate_button)

        # Scrollable widget
        self.scroll_area_widget = QWidget()
        self.scroll_area_layout = QVBoxLayout(self.scroll_area_widget)
        self.scroll_area_widget.setLayout(self.scroll_area_layout)

        # Dictionary to store checkboxes
        self.checkboxes = {}

        # Number of columns for the grid layout of terms
        num_columns = 3

        # Color list for QGroupBox sections
        colors = ["#5C5C5C", "#4B4B4B", "#6E6E6E"]

        # Add sections and checkboxes for each French term
        for idx, (category, terms) in enumerate(categories.items()):
            group_box = QGroupBox(category)
            group_box_layout = QGridLayout()
            group_box.setLayout(group_box_layout)

            # Set background color and style
            color = colors[idx % len(colors)]
            group_box.setStyleSheet(f"""
                background-color: {color}; 
                border: 1px solid black;                
                font-size: 10pt;
                border-radius: 2px;
                padding: 5px;  
                color: white;
            """)

            row = 0
            column = 0

            # Create checkboxes for each term
            for en_term, fr_term in terms.items():
                unique_id = f"{category}_{fr_term}"
                checkbox = QCheckBox(fr_term)
                checkbox.setObjectName(unique_id)
                checkbox.setStyleSheet("color: white;")
                group_box_layout.addWidget(checkbox, row, column)
                self.checkboxes[unique_id] = checkbox

                column += 1
                if column >= num_columns:
                    column = 0
                    row += 1

            self.scroll_area_layout.addWidget(group_box)

        # Create scroll area and add it to the main layout
        self.scroll_area = QScrollArea()
        self.scroll_area.setWidgetResizable(True)
        self.scroll_area.setWidget(self.scroll_area_widget)
        self.main_layout.addWidget(self.scroll_area)

        # Button to generate the Excel file
        self.generate_button = QPushButton("Générer le fichier Excel")
        self.generate_button.setStyleSheet("background-color: #333; color: white;")
        self.generate_button.clicked.connect(self.on_generate_file)
        self.main_layout.addWidget(self.generate_button)

    def on_validate_description(self):
        # Reset all checkboxes before applying new selection
        for checkbox in self.checkboxes.values():
            checkbox.setChecked(False)

        description = self.description_input.text().strip().lower()

        if description:
            # Split the description into keywords
            keywords = description.split()
            matched_terms = []
            
            # Iterate through all checkboxes and match based on French terms
            for unique_id, checkbox in self.checkboxes.items():
                fr_term = unique_id.split('_', 1)[1].lower()

                matched = False

                # Step 1: Exact match in French or via manual mappings
                for keyword in keywords:
                    if keyword == fr_term or (keyword in manual_mappings and manual_mappings[keyword].lower() == fr_term):
                        checkbox.setChecked(True)
                        matched_terms.append(fr_term)
                        matched = True
                        break

                # Step 2: Partial matching keywords in the description (French only)
                if not matched:
                    for keyword in keywords:
                        if keyword in fr_term:
                            checkbox.setChecked(True)
                            matched_terms.append(fr_term)
                            matched = True
                            break

            # Deduplicate and convert matched French terms to English output
            if matched_terms:
                # Get unique English terms
                unique_en_terms = list(set(terms_fr_to_en[term] for term in matched_terms))
                result = ", ".join(unique_en_terms)  # Output in English
                QMessageBox.information(self, "Termes trouvés", f"Les termes suivants ont été cochés : {result}")
            else:
                QMessageBox.information(self, "Aucun terme trouvé", "Aucun terme ne correspond à la description.")
        else:
            QMessageBox.warning(self, "Description vide", "Veuillez entrer une description.")

    # Make sure the on_generate_file method is inside the MainWindow class
    def on_generate_file(self):
        selected_terms = []

        # Get selected terms from checkboxes
        for unique_id, checkbox in self.checkboxes.items():
            if checkbox.isChecked():
                fr_term = unique_id.split('_', 1)[1].lower()
                english_term = terms_fr_to_en.get(fr_term, "Inconnu")
                selected_terms.append(english_term)

        # Deduplicate selected English terms
        selected_terms = list(set(selected_terms))

        if selected_terms:
            result = ", ".join(selected_terms)
            clipboard = QApplication.clipboard()
            clipboard.setText(result)
            search_and_save(selected_terms)
        else:
            QMessageBox.warning(self, "Aucun terme sélectionné", "Veuillez sélectionner au moins un terme.")

# Function to search and save selected terms into Excel
def search_and_save(terms):
    folder_selected = QFileDialog.getExistingDirectory(None, "Sélectionnez le dossier contenant les fichiers .xlsx")
    
    if not folder_selected:
        return
    
    all_results = []

    for root, dirs, files in os.walk(folder_selected):
        for file in files:
            if file.endswith('.xlsx'):
                file_path = os.path.join(root, file)
                try:
                    df = pd.read_excel(file_path, header=None)
                    
                    for term in terms:
                        result = df[df.apply(lambda row: row.astype(str).str.contains(term, case=False).any(), axis=1)]
                        if not result.empty:
                            all_results.append(result)
                except Exception as e:
                    print(f"Erreur lors de la lecture du fichier {file_path}: {e}")
    
    if all_results:
        final_df = pd.concat(all_results, ignore_index=True)
        save_path = QFileDialog.getSaveFileName(None, "Sauvegarder le fichier", "", "Excel files (*.xlsx)")[0]
        if save_path:
            final_df.to_excel(save_path, index=False, header=False)
            QMessageBox.information(None, "Succès", "Les résultats ont été sauvegardés avec succès.")
        else:
            QMessageBox.warning(None, "Annulé", "La sauvegarde du fichier a été annulée.")
    else:
        QMessageBox.information(None, "Aucun résultat", "Aucun terme correspondant n'a été trouvé.")

if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())
