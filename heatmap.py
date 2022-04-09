# draw a heatmap for the commute time
# read data from excel file
from turtle import fillcolor
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
    return heat_layer


# show commute time as heatmap
def commute_heatmap():

    APIKEY = "pk.eyJ1IjoidmFsZW50aW5mcmxjaCIsImEiOiJjazk4ZzQ3MHcwNmJqM3FybzJxMXh1d2U1In0.aGuJSQwT1ub9friABW4HpQ"
    color = "#D1F0FF"

    # connect every lat, long to (47.31783342759546, 8.795774929882974)
    # create a map
    map = folium.Map(location=[47.31783342759546,
                     8.795774929882974], zoom_start=12, tiles='https://api.mapbox.com/styles/v1/valentinfrlch/cl1rnamlr000214t8kpkn0mev/tiles/256/{z}/{x}/{y}@2x?access_token=' + APIKEY, attr="mapbox", name="Map")


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
    circles = []
    polygons = []
    for key in durations:
        if key in [5, 10, 15, 20, 25, 30, 35, 40, 45, 50, 55, 60, 65, 70, 75, 80, 85, 90, 95, 100]:
            # if lenght is 1 draw a circle with distance as the radius and 47.31783342759546, 8.795774929882974 as the center
            if len(durations[key]) == 1:
                circles.append([key, folium.Circle(location=(47.31783342759546, 8.795774929882974), radius=distance(
                    durations[key][0][0], durations[key][0][1], 47.31783342759546, 8.795774929882974),
                    color=color, opacity=0.1, fill=True, fillcolor=color, fill_opacity=0.1,
                    tooltip="commute time: " + str(key) + "min").add_to(map, name="commute time: " + str(key) + "min")])
            elif len(durations[key]) == 2:
                # use the average of the two lat, long as radius
                radius = (distance(durations[key][0][0], durations[key][0][1], durations[key][1][0], durations[key][1][1]) + distance(
                    durations[key][0][0], durations[key][0][1], 47.31783342759546, 8.795774929882974)) / 2 * 1000
                circles.append([key, folium.Circle(
                    location=(47.31783342759546, 8.795774929882974), radius=radius, color=color,
                    opacity=0.1, fill=True, fillcolor=color, fill_opacity=0.1, tooltip="commute time: " + str(key) + "min")])
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
                polygons.append([key, folium.Polygon(
                    locations=points, color=color, tooltip="commute time: " + str(key) + "min", width=1, opacity=0.3, fill=True, fillcolor=color, fill_opacity=0.1)])

    # save the map
    heatmap().add_to(map, name="heatmap")
    #reverse for loop
    for i in range(len(circles) - 1, -1, -1):
        circles[i][1].add_to(map)
    for i in range(len(polygons) - 1, -1, -1):
        polygons[i][1].add_to(map)
    folium.LayerControl(collapsed=False).add_to(map)
    map.save("index.html")
    # open the map in the browser
    webbrowser.open("index.html")


commute_heatmap()
