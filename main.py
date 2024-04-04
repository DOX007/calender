from kivy.app import App
from kivy.uix.gridlayout import GridLayout
from kivy.uix.button import Button
from kivy.core.window import Window
from kivy.uix.popup import Popup
from kivy.uix.textinput import TextInput
from kivy.uix.dropdown import DropDown
import calendar
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from datetime import datetime

# Define the necessary variables
SPREADSHEET_ID = '1xQV_e_kmp9-ECTVX-9pAP4u4F1ug0JFD58vCRWgc5vE'
API_KEY = 'AIzaSyBmhdog4-Uniir8KGka-A0kVi36au2QyAk'

# Authorize using credentials
scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
credentials = ServiceAccountCredentials.from_json_keyfile_name('credentials.json', scope)
client = gspread.authorize(credentials)

# Access the spreadsheet
spreadsheet = client.open_by_key(SPREADSHEET_ID)

class CalendarApp(App):
    def build(self):
        self.year = 2024  # Initial year
        self.month = 1    # Initial month
        self.cal = calendar.monthcalendar(self.year, self.month)
        self.today = datetime.now().day  # Get today's day
        
        layout = GridLayout(cols=7, padding=1, spacing=1)
        self.update_calendar(layout)
        return layout
    
    def update_calendar(self, layout):
        layout.clear_widgets()
        days = ['Mo', 'Ti', 'On', 'To', 'Fr', 'Lö', 'Sö']

        # Add day labels
        for day in days:
            label = Button(text=day, background_color='#1c99e0', font_size=60)
            layout.add_widget(label)

        # Add day buttons
        for week in self.cal:
            for day in week:
                if day == 0:
                    layout.add_widget(Button(text='', background_color=(0.09, 0.1, 0.15, 0)))
                else:
                    btn = Button(text=str(day), background_color='#171a27', font_size=50)
                    if self.is_date(day):
                        btn.border = (1, 1, 1, 1)
                        btn.border_width = (0.5, 0.5)
                    if day == self.today:  # Highlight today's date
                        btn.background_color='#c86fff'  # Change to the specified color
                    btn.bind(on_press=lambda instance, day=day: self.show_popup(day))
                    layout.add_widget(btn)
                    
    def is_date(self, day):
        # Check if a given day is within the month's days
        if day in range(1, calendar.monthrange(self.year, self.month)[1] + 1):
            return True
        return False
        
    def show_popup(self, day):
        self.popup = Popup(title='Datum den: ' + str(day),
                           size_hint=(None, None), size=(700, 400),
                           auto_dismiss=False)
        content = GridLayout(cols=1, padding=10, spacing=10)
        
        # Text input for number
        self.number_input = TextInput(hint_text='Skriv in dina timmar ...', font_size=30)
        content.add_widget(self.number_input)
        
        # Dropdown list
        dropdown = DropDown()
        for option in ['Normal arbetstid', 'Mertidstimmar', 'Storhelgstimmar']:
            btn = Button(text=option, size_hint_y=None, height=110, background_color='#ffffff', font_size=30)
            btn.bind(on_release=lambda btn: dropdown.select(btn.text))
            dropdown.add_widget(btn)
        
        main_button = Button(text='Arbetstider', size_hint=(None, None), size=(325, 80), background_color='#a2a3ac', font_size=30)
        main_button.bind(on_release=dropdown.open)
        dropdown.bind(on_select=lambda instance, x: setattr(main_button, 'text', x))
        
        # Create a nested GridLayout for buttons
        buttons_layout = GridLayout(cols=2, spacing=10)
        
        # Add the main button and the link button
        buttons_layout.add_widget(main_button)
        
        # Add text link to open popup with text "Lamps are good to have"
        link_button = Button(text='Revidera inlagda tider', size_hint=(None, None), size=(320, 80), background_color='#a2a3ac', font_size=25)
        link_button.bind(on_release=lambda instance: self.show_text_popup())
        buttons_layout.add_widget(link_button)
        
        content.add_widget(buttons_layout)
        
        # Buttons layout
        buttons_layout = GridLayout(cols=2, spacing=1)
        
        # OK button
        ok_button = Button(text='OK', background_color='#171a27', font_size=40)
        ok_button.bind(on_press=lambda instance: self.add_to_google_sheets(day, main_button.text))
        buttons_layout.add_widget(ok_button)
        
        # Cancel button
        cancel_button = Button(text='Avbryt', background_color='#171a27', font_size=40)
        cancel_button.bind(on_press=self.cancel_popup)
        buttons_layout.add_widget(cancel_button)
        
        content.add_widget(buttons_layout)
        
        self.popup.content = content
        self.popup.open()
    
    def show_text_popup(self):
        text_popup = Popup(title='Information', size_hint=(None, None), size=(600, 400), auto_dismiss=True)
        text_content = GridLayout(cols=1, padding=10, spacing=10)

        # Text input for information
        text_input = TextInput(text="?????????????????????????.", readonly=True, font_size=35)
        text_content.add_widget(text_input)

        # Create a button to clear the calendar data
        clear_button = Button(text="Clear my calendar", size_hint=(None, None), size=(300, 50), font_size=25)
        clear_button.bind(on_release=self.clear_calendar_data)
        text_content.add_widget(clear_button)

        text_popup.content = text_content
        text_popup.open()

    def cancel_popup(self, instance):
        self.popup.dismiss()
        
    def add_to_google_sheets(self, day, column):
        try:
            number = self.number_input.text.replace(',', '.')  # Replace comma with dot for decimal values
            number = float(number)  # Convert to float
            
            # Check if a choice has been made in the dropdown menu
            if column == 'Arbetstider':
                # Show error message if no choice has been made
                self.number_input.text = 'Byt typ av arbetstider !!! ...'
                return

            # Calculate the row number based on the day selected
            row = 8 + day
            
            # Determine the column index based on the selected option
            if column == 'Normal arbetstid':
                column_index = 4  # Column D
            elif column == 'Mertidstimmar':
                column_index = 5  # Column E
            elif column == 'Storhelgstimmar':
                column_index = 6  # Column F
            
            # Select the first worksheet
            worksheet = spreadsheet.get_worksheet(0)

            # Write the number to the corresponding cell in the selected column
            worksheet.update_cell(row, column_index, number)

            self.popup.dismiss()
        except ValueError:
            # If input is not a valid number, show an error
            self.number_input.text = 'Fel ... Prova igen.'

    def clear_calendar_data(self, instance):
        try:
            # Select the first worksheet
            worksheet = spreadsheet.get_worksheet(0)
            
            # Clear data in the specified range  worksheet.batch_clear(['D9:G39'])
            worksheet.batch_clear(['D9:G39']) 
            
            # Close the popup after clearing the data
            instance.parent.parent.dismiss()  # Dismiss the popup (2 parents up)
        except Exception as e:
            print("Error clearing data:", e)

# Run the application
if __name__ == '__main__':
    Window.clearcolor = (0.09, 0.1, 0.15, 1)  # Set background color
    CalendarApp().run()
