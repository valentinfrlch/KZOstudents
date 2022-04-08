# read excel file
from cv2 import threshold
import openpyxl
import folium
import webbrowser
import math

path = 'dataset/students.xlsx'

# read file
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
        student['email'] = sheet['G' + str(row)].value.strip()
        student["lat"] = sheet['H' + str(row)].value
        student["long"] = sheet['I' + str(row)].value
    except (ValueError, AttributeError) as e:
        print(e)
        pass
    data.append(student)


def main():
    """plot lat and long on map"""
    # import libraries
    import folium
    import webbrowser
    # create map
    m = folium.Map(location=[47.33, 8.78], zoom_start=12)
    for student in data:
        # add marker
        try:
            folium.Marker(
                location=[student["lat"], student["long"]], popup=student["vorname"] + " " + student["name"] + "\n" + student["klasse"]).add_to(m)
        except ValueError:
            pass
    # save map
    m.save('map.html')
    # open map
    webbrowser.open('map.html')


def distance(lat1, lon1, lat2, lon2):
    """calculate distance between two coordinates"""
    # if any input is None -> return None
    if lat1 is None or lon1 is None or lat2 is None or lon2 is None:
        return None
    else:
        R = 6371
        dLat = math.radians(lat2-lat1)
        dLon = math.radians(lon2-lon1)
        lat1 = math.radians(lat1)
        lat2 = math.radians(lat2)
        a = math.sin(dLat/2) * math.sin(dLat/2) + math.sin(dLon/2) * \
            math.sin(dLon/2) * math.cos(lat1) * math.cos(lat2)
        c = 2 * math.atan2(math.sqrt(a), math.sqrt(1-a))
        d = R * c
        return d


def group(threshold=5):
    """group markers if they are close"""
    m = folium.Map(location=[47.33, 8.78], zoom_start=12)
    #calculate distance for each student
    processed = []
    for student in data:
        average = []
        processed.append(student)
        for other in data:
            # add progress bar in %
            print(
                "\r" + str(round(100*(data.index(student)+1)/len(data), 2)) + "%", end="")
            # ignore if already in group
            if student != other and other not in processed:
                d = distance(student["lat"], student["long"],
                             other["lat"], other["long"])
                # if d is None  ->  no lat or long
                if d is not None and d < threshold:
                    #add to processed
                    processed.append(other)
                    # add other lat and long to average
                    average.append((other["lat"], other["long"]))
                # get average lat of average
                if len(average) > 0:
                    avglat = sum(x[0] for x in average) / len(average)
                    # get average long of average
                    avglong = sum(x[1] for x in average) / len(average)
                    folium.Circle((avglat, avglong), radius=threshold/2, color='#3186cc',
                                  fill=True, fill_color='#3186cc').add_to(m)
    print("\r saving map...")
    m.save('local_groups.html')
    # open map
    webbrowser.open('local_groups.html')


group()
