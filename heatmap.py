# draw a heatmap for the commute time
# read data from excel file
import folium
import shapely
from shapely.geometry.polygon import Polygon
from folium import plugins
from folium.plugins import HeatMap
import matplotlib.pyplot as plt
import webbrowser
import openpyxl
from visualize import distance

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
    map = folium.Map(location=[47.38, 8.53], zoom_start=12, blur=70)
    # create a heatmap layer
    # ignore NoneType
    # get lat, long and commute if not None
    d = []
    for student in data:
        if student["lat"] is not None and student["long"] is not None:
            d.append([student["lat"], student["long"]])

    heat_layer = HeatMap(data=d, blur=100, radius=50)
    # add the heatmap layer to the map
    map.add_child(heat_layer)
    # save the map
    map.save("index.html")
    # open the map in the browser
    #webbrowser.open("heatmap.html")


# show commute time as heatmap
def commute_heatmap():
    # connect every lat, long to (47.31783342759546, 8.795774929882974)
    # create a map
    map = folium.Map(location=[47.31783342759546,
                     8.795774929882974], zoom_start=12)

    #get all commute times
    durations = {}
    commute_times = [student["commute"]
                     for student in data if student["commute"] is not None]
    time = list(set(commute_times))
    # add elements from time to durations
    for i in time:
        durations[i] = []

    # add the commute time for each student to the durations dict
    for student in data:
        if student["commute"] is not None:
            durations[student["commute"]].append(
                (student["lat"], student["long"]))

    # for every key in durations draw a line between the values
    for key in durations:
        # get average distance of the values to 47.31783342759546, 8.795774929882974
        dist = []
        for coords in durations[key]:
            dist.append(
                distance(coords, (47.31783342759546, 8.795774929882974)))
        avg_dist = sum(dist) / len(dist)
        # draw a circle around 47.31783342759546, 8.795774929882974 with avg_dist as radius
        circle = folium.Circle(
            radius=avg_dist,
            location=[47.31783342759546, 8.795774929882974],
            color='#3186cc',
            fill=False).add_to(map)

    #save the map
    map.save("index.html")
    # open the map in the browser
    webbrowser.open("index.html")


commute_heatmap()
