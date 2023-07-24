import PySimpleGUI as ui
from name_manager import name_manager
from ppe_scanner import ppe_scanner

ui.theme('DarkGrey12')

ui.theme_text_color('#FFFFFF')
ui.theme_background_color('#202123')
ui.theme_button_color(('#FFFFFF', '#202123'))
ui.theme_element_text_color('#FFFFFF')
ui.theme_input_background_color('#202123')

ui.set_options(font=("Helvetica", 14))

layout = [
    [ui.Button("Name Manager", key='-MANAGE-', size=(20, 1), font=('Helvetica', 14, 'bold'))],
    [ui.Button("PPE Scanner", key='-SCAN-', size=(20, 1), font=('Helvetica', 14, 'bold'))],
    [ui.Button("Exit", key='-EXIT-', size=(20, 1), font=('Helvetica', 14, 'bold'))]
]

window = ui.Window("Main Menu", layout, finalize=True)

while True:
    event, _ = window.read()

    if event == ui.WIN_CLOSED or event == '-EXIT-':
        break

    if event == '-MANAGE-':
        window.Hide()  # Hide the main menu window
        name_manager(window)  # Pass the main window to name_manager function

    elif event == '-SCAN-':
        window.Hide()  # Hide the main menu window
        ppe_scanner(window)  # Pass the main window to ppe_scanner function

window.close()