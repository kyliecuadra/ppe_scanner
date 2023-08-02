import PySimpleGUI as ui
import csv
import os
import datetime
from ultralytics import YOLO
import cv2
import cvzone
import math
import pandas as pd
import threading
import time

# Set PySimpleGUI theme
ui.theme_text_color('#FFFFFF')
ui.theme_background_color('#202123')
ui.theme_button_color(('#FFFFFF', '#202123'))
ui.theme_element_text_color('#FFFFFF')
ui.theme_input_background_color('#202123')

# Constants for data file and Excel folder
DATA_FILE = 'data.csv'
EXCEL_FOLDER = os.path.join(os.path.expanduser('~'), 'Documents', 'PPE Scores')

# Function to read data from CSV file
def read_data_from_csv():
    data = []
    with open(DATA_FILE, 'r') as file:
        reader = csv.reader(file)
        next(reader)  # Skip the header row
        for row in reader:
            data.append(row[0])  # Only add the Student Name to the data list
    return data

# Functions to convert stars to points and vice versa
def stars_to_points(stars):
    return len(stars)

def points_to_stars(points):
    return '★' * points

# Function to get points for a specific student from CSV file
def get_student_points(student_name):
    with open(DATA_FILE, 'r') as file:
        reader = csv.DictReader(file)
        for row in reader:
            if row['Student Name'].lower() == student_name.lower():
                return row['Points']
    return ''

# Function to save data to Excel file
def save_to_excel(student_name, stars):
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
        df = pd.read_excel(file_path, engine='openpyxl')

        # Check if the student name already exists in the Excel file
        if student_name in df['Student Name'].tolist():
            # If the student name exists, update the existing data with the new score
            df.loc[df['Student Name'] == student_name, 'Points'] = stars
        else:
            # If the student name does not exist, add a new row with the student name and score
            new_row = {'Student Name': student_name, 'Points': stars}
            df = df.append(new_row, ignore_index=True)
    else:
        # If the file does not exist, create a new DataFrame with the student name and score
        data = {'Student Name': [student_name], 'Points': [stars]}
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

# Function to create the GUI layout for scanning PPE
def create_scan_layout(data):
    ppe_layout = [
        [ui.Text("PPE", font=('Helvetica', 14, 'bold'), background_color='#202123')],
        [ui.Image(key="-IMAGE-")],
    ]

    right_column_layout = [
        [ui.Text("Student Name", font=('Helvetica', 14, 'bold'), background_color='#202123')],
        [ui.Combo(data, key='-STUDENT-', readonly=True, size=(20, 1), background_color='#444654', text_color='#FFFFFF')],  # Set readonly to True to prevent typing
        [ui.Text("Points", background_color='#202123')],
        [ui.Text("", key='-STARS-', size=(20, 1), background_color='#202123')],
        [ui.Button("Scan", key='-SCAN-')],  # Add Scan button to start scanning
        [ui.Button("Save", key='-SAVE-', disabled=True)],  # Disable the Save button by default
        [ui.Button("Back", key='-BACK-')]
    ]

    layout = [
        [ui.Column(ppe_layout), ui.VSeperator(pad=(0, 0)), ui.Column(right_column_layout)]
    ]
    return layout

# Function to scan the frame for PPE
def scan_frame():
    global detected_ppe  # Declare detected_ppe as a global variable

    # Function to capture frame after 3 seconds
    def capture_frame_after_3_seconds():
        time.sleep(3)  # Wait for 3 seconds
        _, img = cap.read()  # Capture a frame from the webcam
        cv2.imwrite('temporary_frame.png', img)  # Save the frame as a temporary image
        print("Frame captured and saved as temporary_frame.png")

    threading.Thread(target=capture_frame_after_3_seconds).start()
    while True:
        _, img = cap.read()

        results = model(img, stream=True)
        for r in results:
            boxes = r.boxes
            for box in boxes:
                detected_this_frame = set()
                x1, y1, x2, y2 = box.xyxy[0]
                x1, y1, x2, y2 = int(x1), int(y1), int(x2), int(y2)
                w, h = x2 - x1, y2 - y1

                conf = math.ceil((box.conf[0] * 100)) / 100
                cls = int(box.cls[0])
                if classNames[cls] in ["Welding Mask", "Coverall", "Safety Gloves", "Apron", "Safety Shoes"] and conf > 0.90:
                    if classNames[cls] not in detected_ppe:
                        cvzone.putTextRect(img, f'{classNames[cls]} {conf}', (max(0, x1), max(35, y1)), scale=1, thickness=1)
                        cv2.rectangle(img, (x1, y1), (x2, y2), myColor, 3)

                        stars += 1
                        detected_this_frame.add(classNames[cls])

                else:
                    stars = 0
                    detected_ppe.clear()

            detected_ppe |= detected_this_frame

            # Update the stars count based on the number of unique PPE items detected
            stars = len(detected_ppe)

            # Limit the stars count to a maximum of 5
            stars = min(stars, 5)

            # Update the maximum stars count if necessary
            max_stars = max(max_stars, stars)

            # Create a string of stars based on the current count
            total_stars = '★' * max_stars

            print(total_stars)
            window['-STARS-'].update(total_stars)

            # Convert frame to RGB
            frame_rgb = cv2.cvtColor(img, cv2.COLOR_RGB2RGBA)

            # Update sg.Image with the new frame
            window['-IMAGE-'].update(data=cv2.imencode('.png', frame_rgb)[1].tobytes())

        # Enable or disable the Save button based on input values
        name_selected = values['-STUDENT-']
        window['-SAVE-'].update(disabled=False)  # Always enable the Save button

        if event == '-SAVE-' and detected_ppe:
            student_name = values['-STUDENT-']
            save_to_excel(student_name, total_stars)
            stars = 0  # Reset stars count after saving
            detected_ppe.clear()  # Clear detected PPE set after saving
        elif event == '-SAVE-' and not name_selected:
            ui.popup("Please select a student name before saving.", background_color='#202123')

        if event == '-BACK-':
            window.close()  # Close the PPE Scanner window
            main_window.UnHide()  # Unhide the main.py window
            break

# Function to start the PPE scanner GUI
def ppe_scanner(main_window):
    global detected_ppe, stars, max_stars  # Declare the global variables here
    global cap, model, classNames, myColor

    cap = cv2.VideoCapture(0, cv2.CAP_DSHOW)  # Open the webcam
    model = YOLO("best.pt")
    classNames = ["Welding Mask", "Apron", "Coverall", "Safety Gloves", "Safety Shoes"]
    data = read_data_from_csv()
    window = ui.Window("PPE Scanner", create_scan_layout(data))
    stars = 0
    detected_ppe = set()
    max_stars = 0
    myColor = (0, 0, 255)

    while True:
        event, values = window.read(timeout=5)  # Timeout to prevent GUI from freezing

        if event == ui.WIN_CLOSED:
            break

        if event == '-SCAN-':
            threading.Thread(target=scan_frame).start()  # Start scanning in a new thread

    # Release the video capture and close the window
    cap.release()
    window.close()

if __name__ == '__main__':
    ppe_scanner()
