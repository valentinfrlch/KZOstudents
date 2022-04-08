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
    map = folium.Map(location=[47.38, 8.53], zoom_start=12, blur=70)
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


# show commute time as heatmap
def commute_heatmap():
    """create heatmap for commute time"""
    # the longer the commute the more the heatmap will be red
    # the shorter the commute the more the heatmap will be blue
    # create a map
    map = folium.Map(location=[47.31, 8.79], zoom_start=12)
    # create a heatmap layer
    # ignore NoneType
    # get lat, long and commute if not None

    # get max and min commute time
    commute_times = []
    for student in data:
        if student["commute"] is not None:
            commute_times.append(student["commute"])
    max_commute = max(commute_times)
    # get the color for every student
    for student in data:
        if student["lat"] is not None and student["long"] is not None and student["commute"] is not None:
            color = dynamic_color(student["commute"], max_commute)
            folium.CircleMarker(
                [student["lat"], student["long"]],
                radius=10,
                popup=student["name"] + ' ' + student["vorname"],
                color=color,
                fill=True,
                fill_color=color,
                fill_opacity=1).add_to(map)
    # save the map
    map.save("commute_heatmap.html")
    # open the map in the browser
    webbrowser.open("commute_heatmap.html")


def dynamic_color(commute, max_commute):
    """
    calculate color based on delta to max and min commute
    the more the delta the more the color will be white
    """
    def color(percentage):
        """multiply percentage by 255 and round to int"""
        # convert to hex red
        str(hex(int(percentage * 2.55)) + 'ff')[2:]
        return int(percentage * 2.55)
    # calculate percentage of max commute
    percentage = (commute / max_commute) * 100
    # calculate the color based on the percentage
    # the closer to 100 the more the color will be red
    # the closer to 0 the more the color will be white
    return color(percentage)


commute_heatmap()
