# draw a heatmap for the commute time
# read data from excel file
import folium
from folium import plugins
from folium.plugins import HeatMap
import matplotlib.pyplot as plt
import webbrowser
import openpyxl

path = "dataset/students.xlsx"
wb = openpyxl.load_workbook(path)
sheet = wb.get_sheet_by_name('Sheet1')

# load data into list
data = []
for row in range(2, sheet.max_row + 1):
    try:
        student = {}
        student['klasse'] = sheet['A' + str(row)].value.strip()
        student['name'] = sheet['B' + str(row)].value.strip()
        student['vorname'] = sheet['C' + str(row)].value.strip()
        student['addresse'] = sheet['D' +
                                    str(row)].value.strip() + ' ' + sheet['E' + str(row)].value.strip()
        student['telefon'] = sheet['F' + str(row)].value
        student["lat"] = sheet['H' + str(row)].value
        student["long"] = sheet['I' + str(row)].value
        student["commute"] = sheet['J' + str(row)].value
    except (ValueError, AttributeError) as e:
        pass
    data.append(student)

# draw a heatmap for the commute time


def heatmap():
    # create a heatmap for the commute time for every student
    # create a map
    map = folium.Map(location=[47.38, 8.53], zoom_start=12)
    # create a heatmap layer
    # ignore NoneType
    heat_layer = HeatMap(data=[
        [[student for student in data if student["lat"] is not None], [student for student in data if student["long"] is not None], [student for student in data if student["commute"] is not None]] for student in data], radius=10)
    # add the heatmap layer to the map
    map.add_child(heat_layer)
    # save the map
    map.save("heatmap.html")
    # open the map in the browser
    webbrowser.open("heatmap.html")


heatmap()