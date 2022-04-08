# get commuting time for each student
import requests
import openpyxl


def get_station(coordinates):
    if coordinates is None:
        return None
    baseURL = "http://transport.opendata.ch/v1/locations"
    # send request to transport.opendata.ch
    r = requests.get(baseURL, params={
                     "x": coordinates[0], "y": coordinates[1]})
    # get json
    data = r.json()
    try:
        return data["stations"][1]["id"]
    except IndexError:
        return None


def convert_to_minutes(time):
    # split time into hours, minutes and seconds
    hours, minutes, seconds = time.split("d")[1].split(":")
    # convert to minutes
    minutes = int(hours)*60 + int(minutes)
    return minutes


def get_time(station_id):
    baseURL = "http://transport.opendata.ch/v1/connections"
    # send request to transport.opendata.ch
    r = requests.get(baseURL, params={
                     "from": station_id, "to": "Wetzikon, Bahnhof", "limit": "1"})
    # get json
    data = r.json()

    try:
        return convert_to_minutes(data["connections"][0]["duration"])
    except (IndexError, KeyError):
        return None


def mainloop():
    # load excel file
    path = "dataset/students.xlsx"
    wb = openpyxl.load_workbook(path)
    # get sheet
    sheet = wb.get_sheet_by_name('Sheet1')

    #load data into list
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
            pass
        data.append(student)

    # calculate the time for each student
    for student in data:
        # add progresss bar
        print("\r" + str(round(100*(data.index(student)+1)/len(data), 2)) + "%", end="")
        x = student["lat"]
        y = student["long"]
        time = get_time(get_station((x, y)))
        # add time to student
        # write to excel file
        sheet['J' + str(data.index(student) + 2)].value = time
        #save every 100 students
        if data.index(student) % 100 == 0:
            wb.save(path)
    # save excel file
    wb.save(path)


mainloop()
