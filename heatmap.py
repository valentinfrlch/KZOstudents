# draw a heatmap for the commute time
# read data from excel file
import folium
import shapely
import scipy
from shapely.geometry.polygon import Polygon
from folium import plugins
from folium.plugins import HeatMap
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
    # webbrowser.open("heatmap.html")


# show commute time as heatmap
def commute_heatmap():

    color = "red"

    # connect every lat, long to (47.31783342759546, 8.795774929882974)
    # create a map
    map = folium.Map(location=[47.31783342759546,
                     8.795774929882974], zoom_start=12)

    # get all commute times
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
        # if lenght is 1 draw a circle with distance as the radius and 47.31783342759546, 8.795774929882974 as the center
        if len(durations[key]) == 1:
            folium.Circle(location=(47.31783342759546, 8.795774929882974), radius=distance(
                durations[key][0][0], durations[key][0][1], 47.31783342759546, 8.795774929882974), color=color).add_to(map)
        elif len(durations[key]) == 2:
            # use the average of the two lat, long as radius
            radius = (distance(durations[key][0][0], durations[key][0][1], durations[key][1][0], durations[key][1][1]) + distance(
                durations[key][0][0], durations[key][0][1], 47.31783342759546, 8.795774929882974)) / 2 * 1000
            folium.Circle(
                location=(47.31783342759546, 8.795774929882974), radius=radius, color=color).add_to(map)
        else:
            # calculate the convex hull of the points
            hull = shapely.geometry.MultiPoint(durations[key]).convex_hull
            x, y = hull.exterior.coords.xy
            # convert to tuples
            x = list(x)
            y = list(y)
            points = [(x[i], y[i])
                      for i in range(0, len(x))]
            # interpolate the points

            # draw a polygon
            folium.Polygon(
                locations=points, color=color, tooltip=key, width=).add_to(map)

    # save the map
    map.save("index.html")
    # open the map in the browser
    webbrowser.open("index.html")


commute_heatmap()
