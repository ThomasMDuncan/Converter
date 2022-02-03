from tkinter import *
from tkinter import filedialog
from bs4 import BeautifulSoup
import re
from openpyxl import Workbook
import os.path
from os import path
import zipfile as z
from io import BytesIO

class Converter:
    def __init__(self):
        self.__file_name = ""
        self.__extracted_name = ""
        self.__folder_name = None
        self.xlsx_name = ""
        self.chosen = False
        self.extract_kmz = 0

    def set_file_name(self, file_name):
        self.__file_name = file_name
        #print(f"set_file_name: {file_name}")

    def set_folder_name(self, folder_name):
        self.__folder_name = folder_name
        #print(f"set_folder_name: {folder_name}")

    def final_file_name(self):
        xlsx_file_name = ""
        if self.__folder_name == None:
            xlsx_file = self.__file_name.rsplit(".", 1)[0].rsplit("/", 1)[1] + ".xlsx"
            if self.__file_name.rsplit(".", 1)[1] == "zip":
                xlsx_folder = self.__file_name.rsplit(".", 1)[0] + "/"
                xlsx_file_name = xlsx_folder + self.__extracted_name.split(".")[0] + ".xlsx"
            else:
                xlsx_folder = self.__file_name.rsplit("/", 1)[0] + "/"
                xlsx_file_name = xlsx_folder + xlsx_file
        else:
            if self.__file_name == "":
                xlsx_file_name = self. __folder_name + "/" + self.__extracted_name.split(".")[0] + ".xlsx"
            else:
                xlsx_file = self.__file_name.rsplit(".", 1)[0].rsplit("/", 1)[1] + ".xlsx"
                if self.__file_name.rsplit(".", 1)[1] == "zip":
                    xlsx_file_name = self.__folder_name + "/" + self.__file_name.rsplit(".", 1)[0].rsplit("/", 1)[1] + "/" + self.__extracted_name.split(".")[0] + ".xlsx"
                else:
                    xlsx_file_name = self.__folder_name + "/" + xlsx_file
        return xlsx_file_name

    def get_file_name(self):
        return self.__file_name

    def get_folder_name(self):
        return self.__folder_name

    def FileCheck(self, file_name):
        if file_name == "":
            return 1
        try:
            open(file_name, "r")
            return 1
        except IOError:
            print("Error: File does not exist or is unreadable")
            return 0

    def clear(self):
        self.__file_name = ""
        self.__extracted_name = ""
        #self.__folder_name = None
        self.xlsx_name = ""
        #self.chosen = False
        self.extract_kmz = 0

    def convert2(self):
        csv_file_name = ""
        if self.__folder_name == None:
            for i in range(len(self.__file_name) - 1, 0, -1):
                if self.__file_name[i] == '/':
                    csv_file = self.__file_name[:i]
                    fiilepart = self.__file_name[i:].split('.', 1)[0] + ".csv"
                    csv_file_name = csv_file + fiilepart
                    break
        else:
            for i in range(len(self.__file_name) - 1, 0, -1):
                if self.__file_name[i] == '/':
                    fiilepart = self.__file_name[i:].split('.', 1)[0] + ".csv"
                    csv_file_name = self.__folder_name + fiilepart
                    break
    def convert(self):
        """
        Open the KML. Read the KML. Open a CSV file. Process point data to be a CSV row.
        """
        if self.FileCheck(self.__file_name) == 0:
            return
        else:
            kml_list = []
            if self.__file_name == "":
                for file in os.listdir(self.__folder_name):
                    if file.endswith(".kml"):
                        kml_list.append((file, file))
                    elif file.endswith(".kmz"):
                        with z.ZipFile(self.__folder_name + "/" + file, 'r') as zipf:
                            for name2 in zipf.namelist():
                                if name2.split('.')[-1] == 'kml':
                                    zfiledata = BytesIO(zipf.read(name2))
                                    kml_list.append((zfiledata, name2))
                            zipf.close()
            else:
                if self.__file_name.split(".")[-1] == "kml":
                    kml_list.append((open(self.__file_name, 'r'), self.__file_name.rsplit("/", 1)[1]))
                elif self.__file_name.split(".")[-1] == "kmz":
                    with z.ZipFile(self.__file_name, 'r') as zipf:
                        for name2 in zipf.namelist():
                            if name2.split('.')[-1] == 'kml':
                                zfiledata = BytesIO(zipf.read(name2))
                                kml_list.append((zfiledata, name2))
                        zipf.close()
                elif self.__file_name.split(".")[-1] == "zip":
                    if self.__folder_name is not None:
                        os.mkdir(self.__folder_name + "/" + self.__file_name.rsplit(".", 1)[0].rsplit("/", 1)[1])
                    else:
                        os.mkdir(self.__file_name.rsplit(".", 1)[0])
                    #self.__folder_name = self.__file_name.rsplit(".", 1)[0]
                    with z.ZipFile(self.__file_name, 'r') as zfile:
                        for name in zfile.namelist():
                            if name.split('.')[-1] == 'kmz':
                                zfiledata = BytesIO(zfile.read(name))
                                if self.extract_kmz:
                                    this = name.split('.')[0] + '.kmz', self.__file_name.rsplit(".", 1)[0]
                                    zfile.extract(name, self.__file_name.rsplit(".", 1)[0].rsplit("/", 1)[0])
                                with z.ZipFile(zfiledata) as zfile2:
                                    for name2 in zfile2.namelist():
                                        if name2.split('.')[-1] == 'kml':
                                            zfiledata2 = BytesIO(zfile2.read(name2))
                                            kml_list.append((zfiledata2, name2))
                                    zfile2.close()
                        zfile.close()
                else:
                    raise ValueError("NOT SUpported")
            for kml_file in kml_list:
                self.__extracted_name = kml_file[1]
                self.xlsx_name = self.final_file_name()
                #with open(kml_file, 'r') as f:
                # creating the soup object to parse through and then setting that object to the Points folder we want.
                s = BeautifulSoup(kml_file[0], 'lxml')
                header = []
                for i in s.find_all("folder"):
                    folder = s.folder.extract()
                    if folder.find("name").next == "Points":
                        s = folder
                        break
                # point_info_list is a list of dictionaries each dictionary holding the info for each point in the KML
                point_info_list = []
                finalSVlist = []
                # looping through each placemark tak and extracting it, each placemark tag surrounds the info for 1 point.
                for points in s.find_all("placemark"):
                    point = s.placemark.extract()
                    name = point.find("name").next
                    # building temp_header and point_info with some set values from the coordinates tag.
                    temp_header = ['ExactLat', 'ExactLon', 'Altitude']
                    point_info = {}
                    coords = point.find_all("coordinates")[0].next
                    coords = coords.split(',')
                    point_info["ExactLat"] = coords[1].strip('\n')
                    point_info["ExactLon"] = coords[0].strip('\n')
                    point_info["Altitude"] = coords[2].strip('\n')
                    #table1_Stuff = ["Position: ", "1-Sigmas: ", "Ang. Rate: ", "Acceleration: "]
                    table1_Stuff = []
                    table2_Stuff = ["", "1-Sigmas: "]
                    SVs = []
                    SV_table = []
                    satTableKeys = []
                    # looping through each piece of data for each point, which are separeted into rows of a table.
                    # The first column of the row is the name of the piece of data which is added to the temp header.
                    # the second column of the row is that data itself, added to the point_info dictionary
                    colls = []
                    for rows in point.find_all("tr"):
                        row = point.tr.extract()
                        # rows with 2 columns are dealt with as expected the first column is the name the second is that data.
                        # any characters that make up things like km/h or degrees are striped
                        values = row.find_all("td")
                        if "-" in values[0].next:
                            for val in values:
                                table1_Stuff.append(val.next)
                            continue
                        elif values[0].next in {"Lat", "Lon", "Hgt"}:
                            row2 = point.tr.extract()
                            row3 = point.tr.extract()
                            rowL = [row, row2, row3]
                            vals = []
                            for Trow in rowL:
                                values2 = Trow.find_all("td")
                                for val in values2:
                                    if len(val.next) == 1 and ord(val.next) == 160:
                                        values2.remove(val)
                                vals.append(values2)
                            k = 0
                            for thing in table1_Stuff:
                                for j in range(len(vals)):
                                    if thing == "1-Sigmas: ":
                                        temp_header.append(thing + vals[j][2].next)
                                        point_info[thing + vals[j][2].next] = re.sub("[^\d\.\-]", "", vals[j][3].next)
                                        temp_header.append(thing + vals[j][4].next)
                                        point_info[thing + vals[j][4].next] = re.sub("[^\d\.\-]", "", vals[j][5].next)
                                        k = 4
                                    else:
                                        temp_header.append(thing + vals[j][k].next)
                                        point_info[thing + vals[j][k].next] = re.sub("[^\d\.\-]", "", vals[j][k+1].next)
                                k += 2
                            continue
                        elif values[0].next == "Roll":
                            row2 = point.tr.extract()
                            row3 = point.tr.extract()
                            rowL = [row, row2, row3]
                            vals = []
                            for Trow in rowL:
                                vals.append(Trow.find_all("td"))
                            k = 0
                            for thing in table2_Stuff:
                                for j in range(len(vals)):
                                    temp_header.append(thing + vals[j][k].next)
                                    point_info[thing + vals[j][k].next] = re.sub("[^\d\.\-]", "", vals[j][k + 1].next)
                                k += 2
                            continue
                        elif values[0].next == "Track Angle":
                            for i in range(0, num_colls, 2):
                                temp_header.append(values[i].next)
                                point_info[values[i].next] = re.sub("[^\d\.\-]", "", values[i + 1].next)
                            regex_let = re.compile('[^a-zA-Z]')
                            regex_num = re.compile('[^0-9/]')
                            sats = point.tr.extract().find_all("td")
                            sat_num = point.tr.extract().find_all("td")
                            SV_nums = []
                            SV_names = []
                            SVs_Info = []
                            for m in range(1, len(sats)):
                                SV_names.append(regex_let.sub('', sats[m].next))
                                SV_num = regex_num.sub('', sats[m].next).split('/')
                                SVs_Info.append(SV_num)
                                SV_names.append(SV_num[len(SV_num)-1])
                            for p in range(1, len(sat_num)):
                                SV_nums.append(sat_num[p].next)
                            SVs.append(SV_names)
                            SVs.append(SV_nums)
                            SV_dict = {}
                            SV_table_info = []
                            for SV_rows in point.find_all("tr"):
                                SV_row = point.tr.extract()
                                SV = []
                                SVname = SV_row.td.extract().next
                                SV_table_info.append(SVname)
                                for SV_colll in SV_row.find_all("td"):
                                    SV_col = SV_row.td.extract()
                                    SV.append(SV_col.next)
                                SV_dict[SVname] = SV
                                SV_table.append(SV)
                            sat = 0
                            satNum = 1
                            indexer = 0
                            newSVtable = []
                            for sIndex in range(0, len(SV_names), 2):
                                temp_header.append(SV_names[sIndex] + " SVs Used")
                                point_info[SV_names[sIndex] + " SVs Used"] = SVs_Info[sIndex//2][0]
                                temp_header.append(SV_names[sIndex] + " SVs Tracked")
                                point_info[SV_names[sIndex] + " SVs Tracked"] = SVs_Info[sIndex//2][1]
                            svCountIndex = 1
                            svIndexer = SV_names[svCountIndex]
                            satTableKeys.append('Sats')
                            satTableKeys.append('Sat Used')
                            for key in SV_dict.keys():
                                if key == 'Used':
                                    continue
                                satTableKeys.append((key))
                            satTableKeys.append("Sat UTC")
                            for svIndex in range(len(SV_nums)):
                                svIndexer = SV_names[svCountIndex]
                                if svIndex == int(svIndexer):
                                    svCountIndex += 2
                                newSVtableRow = []
                                newSVtableRow.append(SV_names[svCountIndex - 1] + " " + SV_nums[svIndex])
                                newSVtableRow.append(SV_dict['Used'][svIndex])
                                for svRowInfo in SV_dict.keys():
                                    if svRowInfo == 'Used':
                                        continue
                                    newSVtableRow.append(SV_dict[svRowInfo][svIndex])
                                newSVtableRow.append(point_info['UTC'])
                                finalSVlist.append((newSVtableRow))
                            break
                        num_colls = len(values)
                        if num_colls % 2 == 0:
                            for i in range(0, num_colls, 2):
                                if values[i].next == 'SV':
                                    sv_header = temp_header[-1] + ' SV'
                                    temp_header.append(sv_header)
                                    point_info[sv_header] = values[i + 1].next
                                else:
                                    temp_header.append(values[i].next)
                                    if values[i].next in ["UTC", "Type", "Mode", "INS"]:
                                        point_info[values[i].next] = values[i + 1].next
                                    else:
                                        point_info[values[i].next] = re.sub("[^\d\.\-]", "", values[i+1].next)
                        else:
                            satellites = row.find_all("td")[0].next
                            if "Satellites" in satellites:
                                for i in range(1, len(satellites)):
                                    if satellites[i] == '(':
                                        temp_header.append(satellites[:i - 1])
                                        point_info[satellites[:i - 1]] = satellites[i + 1]
                    point_info_list.append(point_info)
                    if len(header) != 0:
                        # finding which header, the previous or current, is longer as it will become the "parent" header
                        if len(header) > len(temp_header):
                            max_header = header
                            min_header = temp_header
                        else:
                            max_header = temp_header
                            min_header = header
                        # loop through the longer header and if the corresponding element in min_header does not match
                        # and max_header doesn't already contain this element insert it into max_header
                        for index in range(len(max_header)):
                            if index < len(min_header) and max_header[index] != min_header[index]:
                                contains = False
                                for i in max_header:
                                    if i == min_header[index]:
                                        contains = True
                                if not contains:
                                    max_header.insert(index, min_header[index])
                        header = max_header
                    else:
                        header = temp_header
                # csv_list is the final list of lists that will be given to the csv writer
                csv_list = []
                # looping through all the point_info dictionaries and each column_head in the header
                # if the point_info dict contains the column_head then you add the corresponding data to the final_row list
                # if the dict does not contain the column_head then add a blank space to the list
                BIheader = ['ExactLat', 'ExactLon', 'Altitude', 'UTC', 'Time', 'Week', 'Type', 'Mode', 'INS', 'PDOP', 'Corr Age', 'Used', 'Track', 'Position:-Lat', 'Position:-Lon',
                            'Position:-Hgt', '1-Sigmas:-East', '1-Sigmas:-North', '1-Sigmas:-Hgt', 'Ang. RateSemi-maj', 'Ang. RateSemi-min', 'Ang. RateOrient.', 'AccelerationX',
                            'AccelerationY', 'AccelerationZ', 'EGM96 Hgt', 'Velocity', 'Roll', 'Pitch', 'Heading', '1-Sigmas: Roll', '1-Sigmas: Pitch', '1-Sigmas: Heading', 'Track Angle',
                            'GPS SVs Used', 'GPS SVs Tracked', 'SBAS SVs Used', 'SBAS SVs Tracked', 'GLONASS SVs Used', 'GLONASS SVs Tracked', 'GALILEO SVs Tracked', 'GALILEO SVs Used']
                for param in BIheader:
                    if not param in header:
                        header.append(param)
                wb = Workbook()
                wb.active.title = "Point Info"
                wb.create_sheet("Sat Table")
                ws = wb.active
                ws.append(header)
                for row in point_info_list:
                    final_row = []
                    for collumn_head in header:
                        if collumn_head in row:
                            final_row.append(row[collumn_head])
                        else:
                            final_row.append(" ")
                    csv_list.append(final_row)
                    ws.append(final_row)
                ws = wb["Sat Table"]
                ws.append(satTableKeys)
                for table in finalSVlist:
                    ws.append(table)
                wb.save(self.xlsx_name)

# if __name__ == '__main__':
#     converter = Converter()
#     converter.convert()

#converter = Converter()

# Function for opening the
# file explorer window
supported = True

def browse_files():
    file_name = filedialog.askopenfilename(initialdir="/", title="Select a File",
                                           filetypes=(("all files", "*.*"), ("KML files", "*.kml*"), ("KMZ files", "*.kmz*"), ("ZIP files", "*.zip*")))
    if len(file_name) > 0:
        if file_name.split(".")[-1] in ["kml", "kmz", "zip"]:
            supported = True
            converter.set_file_name(file_name)
            # Change label contents
            label_file_explorer.configure(text="File: " + file_name)
            if not converter.chosen:
                for i in range(len(file_name) - 1, 0, -1):
                    if file_name[i] == '/':
                        label_folder_explorer.configure(text="Destination: " + file_name[:i])
                        break
        else:
            supported = False
            label_status.configure(text="File Not Supported", fg="red")



def browse_folders():
    # Allow user to select a directory and store it in global var
    # called folder_path
    folder_name = filedialog.askdirectory()
    if len(folder_name) > 0:
        converter.set_folder_name(folder_name)
        converter.chosen = True
        label_folder_explorer.configure(text="Destination: " + folder_name)

def reset_folder():
    if converter.get_file_name() == "":
        converter.set_folder_name("")
        converter.chosen = None
        label_folder_explorer.configure(text="Destination")
    else:
        converter.set_folder_name(converter.get_file_name().rsplit("/")[0])
        converter.chosen = None
        label_folder_explorer.configure(text= "Destination: " + label_file_explorer.cget("text").rsplit(" ", 1)[1].rsplit("/", 1)[0])

def xlsx_path():
    #converter.xlsx_name = converter.final_file_name()
    if converter.get_file_name().rsplit(".", 1)[1] != "zip":
        return path.exists(converter.final_file_name())
    return False

def convert():
    if not supported:
        label_status.configure(text="FILE NOT SUPPORTED", fg="red")
        return
    elif converter.get_file_name() == "":
        if converter.get_folder_name() == None:
            label_status.configure(text="Choose a File", fg="orange")
        else:
            try:
                converter.convert()
                converter.clear()
                label_status.configure(text="Conversion Successful, file location: \n" + converter.xlsx_name,
                                       fg="green")
            except:
                label_status.configure(text="Error in conversion", fg="red")
                # label_file_explorer.configure(text="Error occurred in conversion, email me")
    elif xlsx_path():
        label_status.configure(text="File already exists in chosen directory", fg="red")
    else:
        label_status.configure(text="Status: Converting...", fg="blue")
        window.update_idletasks()
        if converter.get_file_name().split(".")[-1] in ["kml", "kmz", "zip"]:
            try:
                converter.convert()
                converter.clear()
                label_status.configure(text="Conversion Successful, file location: \n" + converter.xlsx_name,
                                       fg="green")
            except:
                label_status.configure(text="Error in conversion", fg="red")
                # label_file_explorer.configure(text="Error occurred in conversion, email me")
        else:
            label_status.configure(text="File not Compatable", fg="red")


#Create the root window
converter = Converter()

window = Tk()

var = IntVar()

def sel():
    converter.check_var = var.get()

def quit():
    window.destroy()

# Set window title
window.title('.kml to .csv converter')

# Set window size
window.geometry("700x500")

# Set window background color
window.config(background="white")

browse_frame = Frame(window, width=700, height=25)
browse_frame.pack(side="top", fill="x")

buttons_frame = Frame(window, width=700, height=25, pady="25")
buttons_frame.pack(side="top", fill="x")

label_frame = Frame(window, width=700, height=100, pady=25)
label_frame.pack(side="top", fill="x")


#clear = Button(btns_frame, text = "C", fg = "black", width = 32, height = 3, bd = 0, bg = "#eee", cursor = "hand2", command = lambda: btn_clear()).grid(row = 0, column = 0, columnspan = 3, padx = 1, pady = 1)

# Create a File Explorer label
label_file_explorer = Label(browse_frame, text="No File Chosen", width=52, height=2, fg="blue")#.grid(row=0, column=1, columnspan=3, padx=1, pady=1)

button_explore = Button(browse_frame, text="Browse Files", width=13, command=browse_files)#.grid(row=0, column=0, padx=15, pady=1)

label_folder_explorer = Label(browse_frame, text="Destination", width=52, height=2, fg="blue")#.grid(row=1, column=1, columnspan=3, padx=1, pady=1)

button_folder_explore = Button(browse_frame, text="Browse Folders", width=13, command=browse_folders)#.grid(row=1, column=0, padx=15, pady=1)

button_folder_reset = Button(browse_frame, text="Reset", command=reset_folder)#.grid(row=1, column=3, padx=1, pady=1)

check_xlsxAndkmz = Checkbutton(buttons_frame, text="xlsx and kmz", width=9, variable=var, onvalue=1, offvalue=0, command=sel)     #.grid(row=0, column=0)

button_run = Button(buttons_frame, text="Convert", width=13, command=convert, justify=LEFT)                #.grid(row=0, column=1)

button_exit = Button(buttons_frame, text="Exit", width=13, command=quit, justify=RIGHT)                     #.grid(row=0, column=2)

label_status = Label(label_frame, text="Status: waiting for file", width=20, height=4, fg="blue", bg="white")         #.grid(row=1, column=0)

# -----------------------

label_file_explorer.grid(row=0, column=5, columnspan=3, padx=1, pady=1)

button_explore.grid(row=0, column=0, padx=15, pady=1)

label_folder_explorer.grid(row=1, column=5, columnspan=3, padx=1, pady=1)

button_folder_explore.grid(row=1, column=0, padx=15, pady=1)

button_folder_reset.grid(row=1, column=500, padx=1, pady=1)

check_xlsxAndkmz.grid(row=0, column=0, padx=15, pady=1)

button_run.grid(row=0, column=1, padx=180, pady=1)

button_exit.grid(row=0, column=2, padx=1, pady=1)

# check_xlsxAndkmz.pack(side="left")
# button_run.pack(side="left")
# button_exit.pack()
label_status.grid(row=0, column=0)
window.mainloop()