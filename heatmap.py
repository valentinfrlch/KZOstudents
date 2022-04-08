# draw a heatmap for the commute time
# read data from excel file
import folium
import matplotlib.pyplot as plt
import webbrowser
from re import T
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
    # connect students with same commute time
    m = folium.Map(location=[47.33, 8.78], zoom_start=12)
    for student in data:
        # connect students with same commute time with a polygon
        try:
            folium.Polygon(
                locations=[
                    [student["lat"], student["long"]],
                    [student["lat"], student["long"] + student["commute"]],
                    [student["lat"] + student["commute"],
                        student["long"] + student["commute"]],
                    [student["lat"] + student["commute"], student["long"]]],
                color='#3186cc',
                fill=False).add_to(m)
        except (ValueError, TypeError):
            pass

            # save map
            m.save('map.html')
            # open map
            webbrowser.open('map.html')


heatmap()
