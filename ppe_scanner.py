import PySimpleGUI as ui
import csv
import os
import datetime
import pandas as pd

ui.theme_text_color('#FFFFFF')               # White text
ui.theme_background_color('#202123')         # Background color #343541
ui.theme_button_color(('#FFFFFF', '#202123')) # Buttons: Background #343541, Text #FFFFFF
ui.theme_element_text_color('#FFFFFF')        # Element text color #FFFFFF
ui.theme_input_background_color('#202123')    # Input element background color #FFFFFF


DATA_FILE = 'data.csv'
EXCEL_FOLDER = os.path.join(os.path.expanduser('~'), 'Documents', 'PPE Scores')

def read_data_from_csv():
    data = []
    with open(DATA_FILE, 'r') as file:
        reader = csv.reader(file)
        next(reader)  # Skip the header row
        for row in reader:
            data.append(row[0])  # Only add the Employee Name to the data list
    return data

def stars_to_points(stars):
    return len(stars)

def points_to_stars(points):
    return '★' * points

def get_employee_points(employee_name):
    with open(DATA_FILE, 'r') as file:
        reader = csv.DictReader(file)
        for row in reader:
            if row['Employee Name'].lower() == employee_name.lower():
                return row['Points']
    return ''

def save_to_excel(employee_name, stars):
    # Convert stars to points
    points = stars_to_points(stars)
    stars = points_to_stars(points)
    # Create the Excel folder if it doesn't exist
    os.makedirs(EXCEL_FOLDER, exist_ok=True)

    # Get the current date as the filename in the format "Month Day, Year"
    current_date = datetime.datetime.now().strftime('%B %d, %Y')

    # Check if the Excel file with the same date exists
    file_path = os.path.join(EXCEL_FOLDER, f"{current_date}.xlsx")
    if os.path.exists(file_path):
        # If the file exists, load the existing data
        df = pd.read_excel(file_path)

        # Check if the employee name already exists in the Excel file
        if employee_name in df['Employee Name'].tolist():
            # If the employee name exists, update the existing data with the new score
            df.loc[df['Employee Name'] == employee_name, 'Points'] = stars
        else:
            # If the employee name does not exist, add a new row with the employee name and score
            new_row = {'Employee Name': employee_name, 'Points': stars}
            df = df.append(new_row, ignore_index=True)
    else:
        # If the file does not exist, create a new DataFrame with the employee name and score
        data = {'Employee Name': [employee_name], 'Points': [stars]}
        df = pd.DataFrame(data)

    # Save the DataFrame to the Excel file
    writer = pd.ExcelWriter(file_path, engine='openpyxl')
    df.to_excel(writer, index=False, na_rep='')
    worksheet = writer.sheets['Sheet1']

    # Adjust the column width to fit the content
    for column_cells in worksheet.columns:
        length = max(len(str(cell.value)) for cell in column_cells)
        worksheet.column_dimensions[column_cells[0].column_letter].width = length

    writer.save()


    # Convert points back to stars for displaying in the popup
    ui.popup(f"Data saved to {file_path}\n\nStars: {stars}", background_color='#202123')

def create_scan_layout(data):
    ppe_layout = [
        [ui.Text("PPE", font=('Helvetica', 14, 'bold'), background_color='#202123')],
        [ui.Checkbox("Welding Mask", key='-WELDING_MASK-', enable_events=True, background_color='#202123')],
        [ui.Checkbox("Coverall", key='-COVERALL-', enable_events=True, background_color='#202123')],
        [ui.Checkbox("Apron", key='-APRON-', enable_events=True, background_color='#202123')],
        [ui.Checkbox("Safety Gloves", key='-SAFETY_GLOVES-', enable_events=True, background_color='#202123')],
        [ui.Checkbox("Safety Shoes", key='-SAFETY_SHOES-', enable_events=True, background_color='#202123')]
    ]

    right_column_layout = [
        [ui.Text("Employee Name", font=('Helvetica', 14, 'bold'), background_color='#202123')],
        [ui.Combo(data, key='-EMPLOYEE-', readonly=True, size=(20, 1), background_color='#444654', text_color='#FFFFFF')],  # Set readonly to True to prevent typing
        [ui.Text("Points", background_color='#202123')],
        [ui.Text("", key='-STARS-', size=(20, 1), background_color='#202123')],
        [ui.Button("Save", key='-SAVE-', disabled=True)],  # Disable the Save button by default
        [ui.Button("Back", key='-BACK-')]
    ]

    layout = [
        [ui.Column(ppe_layout), ui.VSeperator(pad=(0, 0)), ui.Column(right_column_layout)]
    ]
    return layout

def ppe_scanner(main_window):
    data = read_data_from_csv()
    window = ui.Window("Scan App", create_scan_layout(data))

    while True:
        event, values = window.read()

        if event == ui.WIN_CLOSED:
            break

        if event in ['-WELDING_MASK-', '-COVERALL-', '-APRON-', '-SAFETY_GLOVES-', '-SAFETY_SHOES-']:
            stars = ''.join('★' if values[key] else '' for key in ['-WELDING_MASK-', '-COVERALL-', '-APRON-', '-SAFETY_GLOVES-', '-SAFETY_SHOES-'])
            window['-STARS-'].update(stars)

        checkboxes = ['-WELDING_MASK-', '-COVERALL-', '-APRON-', '-SAFETY_GLOVES-', '-SAFETY_SHOES-']
        name_selected = values['-EMPLOYEE-']
        save_enabled = any(values[cbx] for cbx in checkboxes) and name_selected
        window['-SAVE-'].update(disabled=False)

        if '-SAVE-' in event and save_enabled:
            employee_name = values['-EMPLOYEE-']
            stars = ''.join('★' if values[key] else '' for key in checkboxes)
            save_to_excel(employee_name, stars)
        elif '-SAVE-' in event and not name_selected:
            ui.popup("Please select an employee name before saving.", background_color='#202123')
        elif '-SAVE-' in event and not any(values[cbx] for cbx in checkboxes):
            ui.popup("Please select at least one PPE item before saving.", background_color='#202123')

        if event == '-BACK-':
            window.close()  # Close the PPE Scanner window
            main_window.UnHide()  # Unhide the main.py window
            break

    window.close()

if __name__ == '__main__':
    ppe_scanner()
