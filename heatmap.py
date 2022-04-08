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
    map = folium.Map(location=[47.38, 8.53], zoom_start=12, blur=50)
    # create a heatmap layer
    # ignore NoneType
    # get lat, long and commute if not None
    d = []
    for student in data:
        if student["lat"] is not None and student["long"] is not None:
            d.append([student["lat"], student["long"]])

    heat_layer = HeatMap(data=d, radius=10)
    # add the heatmap layer to the map
    map.add_child(heat_layer)
    # save the map
    map.save("index.html")
    # open the map in the browser
    #webbrowser.open("heatmap.html")


heatmap()


# show commute time as heatmap
