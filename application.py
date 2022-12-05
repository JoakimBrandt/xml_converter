import tkinter, tkinter.messagebox, customtkinter, re
from tkinter import filedialog
import math, traceback, os.path
import pandas as pd

#customtkinter.set_appearance_mode("Light")  # Modes: "System" (standard), "Dark", "Light"
#customtkinter.set_default_color_theme("blue")  # Themes: "blue" (standard), "green", "dark-blue"

class App(customtkinter.CTk):

    WIDTH = 800
    HEIGHT = 370

    FILE_LOCATION = ""
    SAVE_LOCATION = ""

    def __init__(self):
        super().__init__()

        self.title("Excel-To-XML converter")
        self.geometry(f"{App.WIDTH}x{App.HEIGHT}")
        self.protocol("WM_DELETE_WINDOW", self.on_closing)  # call .on_closing() when app gets closed

        # ============ create two frames ============

        # configure grid layout (2x1)
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)

        self.frame_left = customtkinter.CTkFrame(master=self,
                                                 width=500,
                                                 corner_radius=0)
        self.frame_left.grid(row=0, column=0, sticky="nswe")

        self.frame_right = customtkinter.CTkFrame(master=self)
        self.frame_right.grid(row=0, column=1, sticky="nswe", padx=20, pady=20)

        # ============ frame_left ============

        # configure grid layout (1x11)
        self.frame_left.grid_rowconfigure(0, minsize=10)   # empty row with minsize as spacing
        self.frame_left.grid_rowconfigure(8, minsize=20)    # empty row with minsize as spacing
        self.frame_left.grid_rowconfigure(11, minsize=10)  # empty row with minsize as spacing
            
        self.label_1 = customtkinter.CTkLabel(master=self.frame_left,
                                              text="File selection",
                                              text_font=("Roboto Medium", -16))  # font name and size in px
        self.label_1.grid(row=1, column=0, pady=10, padx=10)

        self.button_select_excel = customtkinter.CTkButton(master=self.frame_left,
                                                text="Upload excel file",
                                                command=self.file_upload)
        self.button_select_excel.grid(row=2, column=0, pady=10, padx=10)

        self.label_selected_file = customtkinter.CTkLabel(master=self.frame_left,
                                                        text = "Selected file: " + self.FILE_LOCATION,
                                                        text_font=("Roboto Medium", 10))
        self.label_selected_file.grid(row=3, column=0, pady=10, padx=10)

        self.button_select_directory = customtkinter.CTkButton(master=self.frame_left,
                                                text="Select save directory",
                                                command=self.save_directory)
        self.button_select_directory.grid(row=4, column=0, pady=10, padx=10)

        self.label_selected_directory = customtkinter.CTkLabel(master=self.frame_left,
                                                        text = "Selected directory: " + self.SAVE_LOCATION,
                                                        text_font=("Roboto Medium", 10))
        self.label_selected_directory.grid(row=5, column=0, pady=10, padx=10)

        self.label_copyright = customtkinter.CTkLabel(master=self.frame_left, 
                                                    text = "© 2022 Joakim Brandt, Finec Analytics AB All Rights Reserved",
                                                    text_font=("Robot Medium", 8))
        self.label_copyright.grid(row=10, column=0, pady=5, padx=5)
        
        # ============ frame_right ============

        # configure grid layout (3x7)
        self.frame_right.rowconfigure((0, 1, 2, 3), weight=1)
        self.frame_right.rowconfigure(7, weight=10)
        self.frame_right.columnconfigure((0, 1), weight=1)
        self.frame_right.columnconfigure(2, weight=0)

        self.frame_info = customtkinter.CTkFrame(master=self.frame_right)
        self.frame_info.grid(row=0, column=0, columnspan=2, rowspan=4, pady=20, padx=20, sticky="nsew")

        # ============ frame_info ============

        # configure grid layout (1x1)
        self.frame_info.rowconfigure(0, weight=1)
        self.frame_info.columnconfigure(0, weight=1)

        self.label_info_1 = customtkinter.CTkLabel(master=self.frame_info,
                                                   text=self.FILE_LOCATION,
                                                   height=100,
                                                   corner_radius=6,  # <- custom corner radius
                                                   fg_color=("white", "gray38"),  # <- custom tuple-color
                                                   justify=tkinter.LEFT)
        self.label_info_1.grid(column=0, row=0, sticky="nwe", padx=15, pady=15)

        self.label_info_2 = customtkinter.CTkLabel(master=self.frame_info,
                                                   text="Upload progress:",
                                                   height=100,
                                                   corner_radius=6,  # <- custom corner radius
                                                   fg_color=("white", "gray38"),  # <- custom tuple-color
                                                   justify=tkinter.LEFT)
        self.label_info_2.grid(column=0, row=0, sticky="nwe", padx=15, pady=5)

        self.progressbar = customtkinter.CTkProgressBar(master=self.frame_info)
        self.progressbar.grid(row=1, column=0, sticky="ew", padx=15, pady=15)
        self.progressbar.set(0)

        self.button_convert = customtkinter.CTkButton(master=self.frame_info,
                                                    text="Convert",
                                                    command=self.convert_file)
        self.button_convert.grid(row=3, column=0, pady=10, padx=10)


    def file_upload(self):
        filename = filedialog.askopenfilename()
        self.FILE_LOCATION = filename
        self.label_selected_file.configure(text="Selected file:\n%s" %re.findall(r"[^\/]+$", self.FILE_LOCATION)[0])
        self.progressbar.set(1)
    
    def save_directory(self):
        directory_name = filedialog.askdirectory()
        self.SAVE_LOCATION = directory_name
        self.label_selected_directory.configure(text="Selected directory:\n%s" %self.SAVE_LOCATION)

    def convert_file(self):

        if (self.FILE_LOCATION == "" or self.SAVE_LOCATION == ""):
            tkinter.messagebox.showerror("Error", "Please select file and save location")
        else:
            dialog = customtkinter.CTkInputDialog(text="Enter a file name:", title="File name")

            file_name = dialog.get_input()

            if (re.search(r'^[a-zA-Z0-9_]*$', file_name)):
                print("converting")
                convert_file(self.FILE_LOCATION, self.SAVE_LOCATION, file_name=file_name)
            else:
                tkinter.messagebox.showerror("Error", "File name can only consists of 0-9, A-Z, a-z and underscores _")

    def change_appearance_mode(self, new_appearance_mode):
        customtkinter.set_appearance_mode(new_appearance_mode)

    def on_closing(self, event=0):
        self.destroy()

class Property():
    def __init__(self, organizational_number = None, apartment_number = None, property_label = None) -> None:
        self.organizational_number = organizational_number,
        self.apartment_number = apartment_number,
        self.property_label = property_label

    def to_json(self):
        temp_dict = {}
        if not math.isnan(self.organizational_number[0]):
            temp_dict["p:BrfOrgNr"] = self.organizational_number[0]
        if not math.isnan(self.apartment_number[0]):
            temp_dict["p:LagenhetsNr"] = self.apartment_number[0] 
        if (self.property_label[0] != 'nan'):
            temp_dict["p:Fastighetsbeteckning"] = self.property_label

        return temp_dict

class Labor():
    def __init__(self, hours, cost, type) -> None:
        self.hours = hours,
        self.cost = cost,
        self.type = type
    
    def to_json(self):
        return {
            "p:AntalTimmar": str(int(self.hours[0])),
            "p:Kostnad": str(int(self.cost[0])),
            "p:TypAvUtfortArbete": self.type
        }

class Case():
    def __init__(self, invoice_number, buyer, property_information: Property, labor: Labor, misc_costs, pay_day, paid_amount, requested_amount) -> None:
        self.requested_amount = requested_amount
        self.pay_day = pay_day
        self.paid_amount = paid_amount
        self.invoice_number = invoice_number
        self.property_information = property_information
        self.labor = labor
        self.buyer = buyer
        self.misc_costs = misc_costs
    
    def to_json(self):
        return {
            "p:BegartBelopp": str(int(self.requested_amount)),
            "p:Betalningsdatum": self.pay_day.date(),
            "p:BetaltBelopp": str(int(self.paid_amount)),
            "p:FakturaNr": self.invoice_number,
            "p:Fastighet": self.property_information.to_json(),
            "p:Kopare": self.buyer,
            "p:OvrigKostnad": str(int(self.misc_costs)),
            "p:UtfortArbete": self.labor.to_json()
        }

def excel_to_dictionary(path_to_excel):
    try:
        df = pd.read_excel(path_to_excel, engine='openpyxl')
    except Exception as e:
        print(e)
    dict = df.to_dict()
    entries = len(dict['ns1:NamnPaBegaran'])
    return [df.to_dict(), entries]

def convert_file(file_location, save_location, file_name):
    import xml.etree.cElementTree as e

    counter = 0
    cases = []

    excel_information = excel_to_dictionary(file_location)
    excel_dict = excel_information[0]
    amount_of_entries = excel_information[1]

    while counter < amount_of_entries:
        cases.append(
            Case(
                excel_dict['ns1:FakturaNr'][counter],
                excel_dict['ns1:Kopare'][counter],
                Property(
                    organizational_number   = excel_dict['ns1:BrfOrgNr'][counter],
                    apartment_number        = excel_dict['ns1:LagenhetsNr'][counter],
                    property_label          = excel_dict['ns1:Fastighetsbeteckning'][counter],
                ),
                Labor(
                    hours   = excel_dict['ns1:AntalTimmar'][counter],
                    cost    = excel_dict['ns1:Kostnad'][counter],
                    type    = excel_dict['ns1:TypAvUtfortArbete'][counter],
                ),
                excel_dict['ns1:OvrigKostnad'][counter],
                excel_dict['ns1:Betalningsdatum'][counter],
                excel_dict['ns1:BetaltBelopp'][counter],
                excel_dict['ns1:BegartBelopp'][counter]
            )
        )
        counter += 1

    data = e.Element('p:Begaran')
    data.set('xmlns:p', 'http://xmls.skatteverket.se/se/skatteverket/skattered/begaran/1.0')
    data.set('xmlns:xsi', 'http://www.w3.org/2001/XMLSchema-instance')

    e.SubElement(data, 'p:NamnPaBegaran').text = 'GrönTeknik'
    e.SubElement(data, 'p:TypAvBegaran').text = 'GRON_TEKNIK'
    e.SubElement(data, 'p:Utforare').text = '5590871355'

    try:
        for case in cases:
            storage_dict = case.to_json()

            temp_case = e.SubElement(data, 'p:Arende')
            e.SubElement(temp_case, 'p:FakturaNr').text = str(storage_dict['p:FakturaNr'])
            e.SubElement(temp_case, 'p:Kopare').text = str(storage_dict['p:Kopare'])

            temp_fastighet = e.SubElement(temp_case, 'p:Fastighet')
            if (storage_dict['p:Fastighet'].get('p:BrfOrgNr') != None): 
                e.SubElement(temp_fastighet, 'p:LagenhetsNr').text = storage_dict['p:Fastighet'].get('p:LagenhetsNr')

            if (storage_dict['p:Fastighet'].get('p:BrfOrgNr') != None): 
                e.SubElement(temp_fastighet, 'p:BrfOrgNr').text = storage_dict['p:Fastighet'].get('p:BrfOrgNr')

            if (storage_dict['p:Fastighet'].get('p:Fastighetsbeteckning') != None): 
                e.SubElement(temp_fastighet, 'p:Fastighetsbeteckning').text = str(storage_dict['p:Fastighet']['p:Fastighetsbeteckning'])

            temp_utfort_arbete = e.SubElement(temp_case, 'p:UtfortArbete')
            e.SubElement(temp_utfort_arbete, 'p:TypAvUtfortArbete').text = str(storage_dict['p:UtfortArbete']['p:TypAvUtfortArbete'])
            e.SubElement(temp_utfort_arbete, 'p:AntalTimmar').text = str(storage_dict['p:UtfortArbete']['p:AntalTimmar'])
            e.SubElement(temp_utfort_arbete, 'p:Kostnad').text = str(storage_dict['p:UtfortArbete']['p:Kostnad'])
            
            e.SubElement(temp_case, 'p:OvrigKostnad').text = str(storage_dict['p:OvrigKostnad'])
            e.SubElement(temp_case, 'p:Betalningsdatum').text = str(storage_dict['p:Betalningsdatum'])
            e.SubElement(temp_case, 'p:BetaltBelopp').text = str(storage_dict['p:BetaltBelopp'])
            e.SubElement(temp_case, 'p:BegartBelopp').text = str(storage_dict['p:BegartBelopp'])
        
        xml = e.tostring(data)

        write_file_name = os.path.join(save_location, file_name+".xml")

        with open(write_file_name, "wb") as f:
            f.write(xml)

    except Exception as e:
        print(e, traceback.format_exc())

if __name__ == "__main__":
    app = App()
    app.mainloop()
