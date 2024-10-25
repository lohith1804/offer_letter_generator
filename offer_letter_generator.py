from kivy.app import App
from kivy.uix.boxlayout import BoxLayout
from kivy.uix.button import Button
from kivy.uix.label import Label
from kivy.uix.filechooser import FileChooserListView
from kivy.uix.popup import Popup
import pandas as pd
from docxtpl import DocxTemplate
import datetime
import os

class OfferLetterGenerator(App):
    def build(self):
        layout = BoxLayout(orientation='vertical', spacing=10)
        self.file_label = Label(text="No file selected")
        layout.add_widget(self.file_label)
        
        file_chooser = FileChooserListView()
        file_chooser.filters = ['*.xls', '*.xlsx']
        layout.add_widget(file_chooser)
        
        upload_button = Button(text="Generate Offer Letters", size_hint=(None, None), size=(200, 50))
        upload_button.bind(on_press=lambda x: self.generate_offer_letters(file_chooser.path, file_chooser.selection))
        layout.add_widget(upload_button)
        
        return layout
    
    def generate_offer_letters(self, path, filename):
        if filename:
            selected_file = filename[0]
            self.file_label.text = f"Selected file: {selected_file}"
            try:
                df = pd.read_excel(selected_file)
                for index, row in df.iterrows():
                    today_date = datetime.datetime.today().strftime('%B %d %y')
                    name = row['NAME']
                    reg_no = row['PIN']
                    college = row['COLLEGE']
                    submission_date = row['SUB_DATE']
                    branch = row['BRANCH']
                    authorizer = row['AUTHORIZER']
                    self.generate_offer_letter(name, reg_no, branch, college, submission_date, authorizer, today_date)
                print("Offer letters generated successfully.")
            except Exception as e:
                print("Error generating offer letters:", e)
        else:
            self.file_label.text = "No file selected"
    
    def generate_offer_letter(self, name, reg_no, branch, college, submission_date, authorizer, today_date):
        folder_path = os.path.join('Offer_letters', college, branch)
        os.makedirs(folder_path, exist_ok=True)
        certificate_filename = f'offerletter_{name}.docx'
        offer_letter_path = os.path.join(folder_path, certificate_filename)
        context = {
        'today_date' : today_date,
        'name' : name,
        'branch' : branch,
        'college' : college,
        'reg_no' : reg_no,
        'authorizer' : authorizer,
        'submission_date' : submission_date
        }
        doc = DocxTemplate("D:\offer_letter_generator\offerletter1.docx")

        doc.render(context)

        doc.save(offer_letter_path)
        popup_layout = BoxLayout(orientation='vertical', padding=10, spacing=10)
        # popup = Popup(title=f'Offer Letter for {name}', content=popup_layout, size_hint=(None, None), size=(400, 300))
        

        # close_button = Button(text='Close')
        # close_button.bind(on_press=popup.dismiss)
        # popup_layout.add_widget(close_button)
        
        # popup.open()

if __name__ == '__main__':
    OfferLetterGenerator().run()

