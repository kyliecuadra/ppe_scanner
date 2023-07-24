import PySimpleGUI as ui
import csv
import os

ui.theme_text_color('#FFFFFF')               # White text
ui.theme_background_color('#202123')         # Background color #343541
ui.theme_button_color(('#FFFFFF', '#202123')) # Buttons: Background #343541, Text #FFFFFF
ui.theme_element_text_color('#FFFFFF')        # Element text color #FFFFFF
ui.theme_input_background_color('#202123')    # Input element background color #FFFFFF


DATA_FILE = 'data.csv'

def read_data_from_csv():
    data = []
    with open(DATA_FILE, 'r') as file:
        reader = csv.reader(file)
        next(reader)  # Skip the header row
        for row in reader:
            data.append(row[0])  # Only add the Employee Name to the data list
    return data

def check_duplicates(new_name):
    with open(DATA_FILE, 'r') as file:
        reader = csv.reader(file)
        for row in reader:
            if row[0].lower() == new_name.lower():
                return True
    return False

def add_data_to_csv(new_name):
    with open(DATA_FILE, 'a', newline='') as file:
        writer = csv.writer(file)
        writer.writerow([new_name])

def delete_data_from_csv(selected_name):
    temp_file = DATA_FILE + '.tmp'
    with open(DATA_FILE, 'r') as file, open(temp_file, 'w', newline='') as temp:
        reader = csv.reader(file)
        writer = csv.writer(temp)
        for row in reader:
            if row[0].lower() != selected_name.lower():
                writer.writerow(row)

    os.remove(DATA_FILE)
    os.rename(temp_file, DATA_FILE)

def create_name_manager_layout(data):
    layout = [
        [ui.Text("Employee Name", font=('Helvetica', 14, 'bold'), background_color='#202123')],
        [ui.Input(key='-NEW_NAME-', background_color='#444654', text_color='#FFFFFF'), ui.Button("Add", key='-ADD-')],
        [ui.Listbox(values=data, size=(30, 6), key='-NAMES-', enable_events=True, background_color='#444654', text_color='#FFFFFF')],
        [ui.Button("Update", key='-UPDATE-'), ui.Button("Delete", key='-DELETE-')],
        [ui.Button("Back to Main Menu", key='-BACK-')]
    ]
    return layout

def name_manager(main_window):
    data = read_data_from_csv()
    window = ui.Window("Name Manager", create_name_manager_layout(data))

    while True:
        event, values = window.read()

        if event == ui.WIN_CLOSED:
            break

        if event == '-ADD-' and values['-NEW_NAME-']:
            new_name = values['-NEW_NAME-'].strip()
            if not check_duplicates(new_name):
                add_data_to_csv(new_name)
                window['-NEW_NAME-'].update('')
                data.append(new_name)
                window['-NAMES-'].update(data)
            else:
                ui.popup(f"{new_name} already exists!", background_color='#202123')

        elif event == '-NAMES-' and values['-NAMES-']:
            selected_name = values['-NAMES-'][0]
            window['-NEW_NAME-'].update(selected_name)

        elif event == '-UPDATE-' and values['-NAMES-']:
            selected_name = values['-NAMES-'][0]
            new_name = values['-NEW_NAME-'].strip()
            if selected_name and new_name and not check_duplicates(new_name):
                delete_data_from_csv(selected_name)
                add_data_to_csv(new_name)
                data.remove(selected_name)
                data.append(new_name)
                window['-NAMES-'].update(data)
                window['-NEW_NAME-'].update('')
            else:
                ui.popup("Please select a name to update and provide a new name (not duplicate).", background_color='#202123')

        elif event == '-DELETE-' and values['-NAMES-']:
            selected_name = values['-NAMES-'][0]
            if ui.popup_yes_no(f"Are you sure you want to delete {selected_name}?", background_color='#202123') == 'Yes':
                delete_data_from_csv(selected_name)
                data.remove(selected_name)
                window['-NAMES-'].update(data)
                window['-NEW_NAME-'].update('')

        elif event == '-BACK-':
            window.close()  # Close the Name Manager window
            main_window.UnHide()  # Unhide the main.py window
            break

    window.close()

if __name__ == '__main__':
    name_manager()