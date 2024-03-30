from kivy.app import App
from kivy.uix.gridlayout import GridLayout
from kivy.uix.button import Button
from kivy.core.window import Window
from kivy.uix.popup import Popup
from kivy.uix.textinput import TextInput
from kivy.uix.dropdown import DropDown
import calendar
import openpyxl

class CalendarApp(App):
    def build(self):
        self.year = 2024  # Initial year
        self.month = 1    # Initial month
        self.cal = calendar.monthcalendar(self.year, self.month)
        
        layout = GridLayout(cols=7, padding=1, spacing=1)
        
        self.update_calendar(layout)
        
        return layout
    
    def update_calendar(self, layout):
        layout.clear_widgets()
        days = ['Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat', 'Sun']

        # Add day labels
        for day in days:
            label = Button(text=day, background_color='#1c99e0', font_size=20)  # Set background color and font size for day labels
            layout.add_widget(label)

        # Add day buttons
        for week in self.cal:
            for day in week:
                if day == 0:
                    layout.add_widget(Button(text='', background_color=(0.09, 0.1, 0.15, 0)))  # Set the same background color for empty cells
                else:
                    btn = Button(text=str(day), background_color='#171a27', font_size=20)  # Set background color and font size for date cells
                    if self.is_date(day):
                        btn.border = (1, 1, 1, 1)  # Set white border
                        btn.border_width = (0.5, 0.5)  # Set border thickness
                    btn.bind(on_press=lambda instance, day=day: self.show_popup(day))
                    layout.add_widget(btn)

    def is_date(self, day):
        # Check if a given day is within the month's days
        if day in range(1, calendar.monthrange(self.year, self.month)[1] + 1):
            return True
        return False
        
    def show_popup(self, day):
        self.popup = Popup(title='Datum den: ' + str(day),
                           size_hint=(None, None), size=(400, 200),
                           auto_dismiss=False)
        content = GridLayout(cols=1, padding=10, spacing=10)
        
        # Text input for number
        self.number_input = TextInput(hint_text='Skriv in dina timmar ...', font_size=20)
        content.add_widget(self.number_input)
        
        # Dropdown list
        dropdown = DropDown()
        for option in ['Normal arbetstid', 'Övertidstimmar', 'Storhelgstillägg']:
            btn = Button(text=option, size_hint_y=None, height=30, background_color='#171a27')
            btn.bind(on_release=lambda btn: dropdown.select(btn.text))
            dropdown.add_widget(btn)
        
        main_button = Button(text='Arbetstider', size_hint=(None, None), size=(150, 30), background_color='#171a27')
        main_button.bind(on_release=dropdown.open)
        dropdown.bind(on_select=lambda instance, x: setattr(main_button, 'text', x))
        
        content.add_widget(main_button)
        
        # Buttons layout
        buttons_layout = GridLayout(cols=2, spacing=1)
        
        # OK button
        ok_button = Button(text='OK', background_color='#171a27', font_size=20)
        ok_button.bind(on_press=lambda instance: self.add_to_excel(day, main_button.text))
        buttons_layout.add_widget(ok_button)
        
        # Cancel button
        cancel_button = Button(text='Avbryt', background_color='#171a27', font_size=20)
        cancel_button.bind(on_press=self.cancel_popup)
        buttons_layout.add_widget(cancel_button)
        
        content.add_widget(buttons_layout)
        
        self.popup.content = content
        self.popup.open()
        
    def cancel_popup(self, instance):
        self.popup.dismiss()
        
    def add_to_excel(self, day, column):
        try:
            number = self.number_input.text.replace(',', '.')  # Replace comma with dot for decimal values
            number = float(number)  # Convert to float
            
            # Calculate the row number based on the day selected
            row = 8 + day
            
            # Determine the column index based on the selected option
            if column == 'Normal arbetstid':
                column_index = 4  # Column D
            elif column == 'Övertidstimmar':
                column_index = 5  # Column E
            elif column == 'Storhelgstillägg':
                column_index = 6  # Column F
            
            # Open the Excel file
            wb = openpyxl.load_workbook('tilmlista.xlsx')
            # Select the second sheet (index 1)
            sheet = wb.worksheets[1]
            # Write the number to the corresponding cell in the selected column
            sheet.cell(row=row, column=column_index).value = number
            # Save the workbook
            wb.save('tilmlista.xlsx')
            self.popup.dismiss()
        except ValueError:
            # If input is not a valid number, show an error
            self.number_input.text = 'Något blev fel, prova igen.'

if __name__ == '__main__':
    Window.clearcolor = (0.09, 0.1, 0.15, 1)  # Set background color
    CalendarApp().run()
