"""
Wafflesian CCA Allocator Algorithm v1.4.0 Beta GUI
==================================================

This program was created for the purposes of simplifying our allocator into a GUI.
"""

from CCAllocator import *
import logging

logging.basicConfig(
    level=logging.DEBUG,
    filename="gui_log.log",
    filemode="a",
    format="%(process)d - %(asctime)s - %(levelname)s - %(message)s",
    datefmt="%d-%b-%y %H:%M:%S",
)

if __name__ == "__main__":
    logging.info("CCAllocator: Program run as main.")
else:
    logging.info("CCAllocator: Program run as module.")

print("Loading tkinter GUI...")
logging.info("GUI: Setting up tkinter.")

# Import tkinter dependencies
try:
    import tkinter
    from tkinter.tix import *
    from tkinter import messagebox
    from tkinter import filedialog
    from tkinter import scrolledtext
except ImportError:
    print(
        "Could not import tkinter dependencies. Make sure you have installed tkinter."
    )

# Initialise
root = tkinter.Tk()

# Get screen dimensions and configure app_frame window dimensions
screen_width = root.winfo_screenwidth() / 1.2
screen_height = screen_width / 1.8
screen_width = str(int(screen_width))
screen_height = str(int(screen_height))
root.geometry(screen_width + "x" + screen_height)
app_frame = root

# Create scrollbar
# app_canvas = tkinter.Canvas(root)
# app_frame = tkinter.Frame(app_canvas)
# app_scrollbar = tkinter.Scrollbar(app_canvas,orient='vertical',command=app_canvas.yview)
# app_canvas.grid(row=1,column=4)
# app_scrollbar.grid(row=1,column=5)
# app_canvas.configure(yscrollcommand=app_scrollbar.set)

# Setup title
root.title("Wafflesian CCAllocator v1.3.0b")

# Logging widget

label_logging = tkinter.Label(app_frame, text="Logging:")

label_logging.grid(row=1, column=4, padx=10, pady=3)

text_logging = scrolledtext.ScrolledText(
    app_frame, width=40, height=30, state="disabled"
)

text_logging.grid(row=2, column=4, rowspan=10, padx=10, pady=3)

# Define button outputs
def output_inputfilepath():
    global PATH
    PATH = entry_inputfilepath.get()
    messagebox.showinfo("Info", "Path: " + str(PATH))
    print(PATH)


def output_inputfilename():
    global FILENAME
    FILENAME = entry_inputfilename.get()
    if FILENAME.endswith(".xlsx"):
        messagebox.showinfo("Info", "File name:" + str(FILENAME))
    else:
        messagebox.showerror(
            "Error",
            "Invalid file type. Make sure your input file ends with .xlsx. Convert the file type if necessary.",
        )
        FILENAME = ""
        print(FILENAME)


def output_inputfilewindow():
    global PATH
    global FILENAME
    file_path = filedialog.askopenfilename()
    if file_path.endswith(".xlsx"):
        FILENAME = file_path.split("/")[-1]
        PATH = file_path[0 : (len(file_path) - len(FILENAME))]
        entry_inputfilepath.delete(0, tkinter.END)
        entry_inputfilepath.insert(0, PATH)
        entry_inputfilename.delete(0, tkinter.END)
        entry_inputfilename.insert(0, FILENAME)
    else:
        messagebox.showerror(
            "Error",
            "Invalid file type. Make sure your input file ends with .xlsx. Convert the file if necessary.",
        )
        PATH = ""
        FILENAME = ""
        print(PATH, FILENAME)


def output_inputccalist():
    global CCALIST
    global CCAALIASES
    CCALIST = entry_inputccalist.get("1.0", "end-1c").splitlines()
    if CCALIST == []:
        messagebox.showerror("Error", "No CCAs have been inputted.")
    else:
        messagebox.showinfo("Info", "List of CCAs have been inputted.")
        ccalist = CCALIST
        for CCA in range(len(ccalist)):
            ccalist[CCA] = CCALIST[CCA].split(" ")
            if len(ccalist[CCA]) > 1:
                CCAALIASES[ccalist[CCA][0]] = []
                for ALIAS in range(1, len(ccalist[CCA])):
                    CCAALIASES[ccalist[CCA][0]].append(ccalist[CCA][ALIAS])
        for CCA in range(len(CCALIST)):
            CCALIST[CCA] = CCALIST[CCA][0]
    print(CCALIST)
    print(CCAALIASES)


def output_inputccatype():
    global CCATYPE
    ccatype = entry_inputccatype.get("1.0", "end-1c").splitlines()
    if ccatype != []:
        for CCA in range(len(ccatype)):
            ccatype[CCA] = ccatype[CCA].split(" ")
            if len(ccatype[CCA]) == 2:
                CCATYPE[ccatype[CCA][0]] = ccatype[CCA][1]
        messagebox.showinfo("Info", "The type of each CCA has been inputted.")
    else:
        CCATYPE = {}
        messagebox.showinfo("Info", "CCA types reset.")
    print(CCATYPE)


def output_sheetorderstudentlist():
    global SHEETORDER
    entry = entry_sheetorderstudentlist.get()
    if entry == "":
        messagebox.showerror("Error", "Nothing in student list sheet position.")
        SHEETORDER["studentList"] = None
    else:
        try:
            int(entry)
        except:
            messagebox.showerror("Error", "Student list sheet position not an integer.")
        else:
            SHEETORDER["studentList"] = int(entry)
            messagebox.showinfo("Info", "Position of student list: " + entry)
    print(SHEETORDER)


def output_sheetorderhealthstats():
    global SHEETORDER
    entry = entry_sheetorderhealthstats.get()
    if entry == "":
        messagebox.showinfo("Info", "Nothing in health stats sheet position.")
        SHEETORDER["healthStats"] = None
    else:
        try:
            int(entry)
        except:
            messagebox.showerror("Error", "Health stats sheet position not an integer.")
        else:
            SHEETORDER["healthStats"] = int(entry)
            messagebox.showinfo("Info", "Position of health stats: " + entry)
    print(SHEETORDER)


def output_sheetordermusic():
    global SHEETORDER
    entry = entry_sheetordermusic.get()
    if entry == "":
        messagebox.showinfo("Info", "Nothing in music sheet position.")
        SHEETORDER["music"] = None
    else:
        try:
            int(entry)
        except:
            messagebox.showerror("Error", "Music sheet position not an integer.")
        else:
            SHEETORDER["music"] = int(entry)
            messagebox.showinfo("Info", "Position of music: " + entry)
    print(SHEETORDER)


def output_sheetorderart():
    global SHEETORDER
    entry = entry_sheetorderart.get()
    if entry == "":
        messagebox.showinfo("Info", "Nothing in art sheet position.")
        SHEETORDER["art"] = None
    else:
        try:
            int(entry)
        except:
            messagebox.showerror("Error", "Art sheet position not an integer.")
        else:
            SHEETORDER["art"] = int(entry)
            messagebox.showinfo("Info", "Position of art: " + entry)
    print(SHEETORDER)


def output_sheetorderspecial():
    global SHEETORDER
    entry = entry_sheetorderspecial.get()
    if entry == "":
        messagebox.showinfo("Info", "Nothing in achievements sheet position.")
        SHEETORDER["special"] = None
    else:
        try:
            int(entry)
        except:
            messagebox.showerror("Error", "Achievements sheet position not an integer.")
        else:
            SHEETORDER["special"] = int(entry)
            messagebox.showinfo("Info", "Position of special achievements: " + entry)
    print(SHEETORDER)


def output_sheetorderccalist():
    global SHEETORDER
    entry = entry_sheetorderccalist.get()
    if entry == "":
        messagebox.showerror("Error", "Nothing in CCA list sheet position.")
        SHEETORDER["CCAList"] = None
    else:
        try:
            int(entry)
        except:
            messagebox.showerror("Error", "CCA list sheet position not an integer.")
        else:
            SHEETORDER["CCAList"] = int(entry)
            messagebox.showinfo("Info", "Position of CCA list: " + entry)
    print(SHEETORDER)


def output_sheetorderchoices():
    global SHEETORDER
    entry = entry_sheetorderchoices.get()
    if entry == "":
        messagebox.showerror("Error", "Nothing in choices sheet position.")
        SHEETORDER["choices"] = None
    else:
        try:
            int(entry)
        except:
            messagebox.showerror("Error", "Choices sheet position not an integer.")
        else:
            SHEETORDER["choices"] = int(entry)
            messagebox.showinfo("Info", "Position of choices: " + entry)
    print(SHEETORDER)


def output_choicesmain():
    global CHOICES
    entry = entry_choicesmain.get()
    if entry == "":
        messagebox.showinfo("Info", "Nothing in main choices.")
    else:
        try:
            int(entry)
        except:
            messagebox.showerror("Error", "Main choices not an integer.")
        else:
            CHOICES["main"] = int(entry)
    print(CHOICES)


def output_choicesother():
    global CHOICES
    entry = entry_choicesother.get()
    if entry == "":
        messagebox.showinfo("Info", "Nothing in other choices.")
    else:
        try:
            int(entry)
        except:
            messagebox.showerror("Error", "Other choices not an integer.")
        else:
            CHOICES["other"] = int(entry)
    print(CHOICES)


def output_resetdata():
    global PATH
    global FILENAME
    global CCALIST
    global CCAALIASES
    global CCATYPE
    global SHEETORDER
    global CHOICES
    global ALOB

    del ALOB
    ALOB = None
    PATH = ""
    FILENAME = ""
    CCALIST = []
    CCAALIASES = {}
    CCATYPE = {}
    SHEETORDER = {
        "studentList": None,
        "healthStats": None,
        "music": None,
        "art": None,
        "special": None,
        "CCAList": None,
        "choices": None,
    }
    CHOICES = {"main": 1, "other": None}
    entry_inputfilepath.delete(0, "end")
    entry_inputfilename.delete(0, "end")
    entry_inputccalist.delete("1.0", "end")
    entry_inputccatype.delete("1.0", "end")
    entry_sheetorderstudentlist.delete(0, "end")
    entry_sheetorderhealthstats.delete(0, "end")
    entry_sheetordermusic.delete(0, "end")
    entry_sheetorderart.delete(0, "end")
    entry_sheetorderspecial.delete(0, "end")
    entry_sheetorderccalist.delete(0, "end")
    entry_sheetorderchoices.delete(0, "end")
    entry_choicesmain.delete(0, "end")
    entry_choicesother.delete(0, "end")
    text_logging.configure(state="normal")
    text_logging.delete(0, "end")
    text_logging.configure(state="disabled")
    print("Data has been reset.")
    logging.info("GUI.output_resetdata: Reset successful.")
    messagebox.showinfo("Info", "Data has been reset.")


def output_startallocation():
    def loginto(errorlevel, text):
        if errorlevel == 0:  # debug
            logging.debug("GUI.output_startallocation: " + text)
        elif errorlevel == 1:  # info
            logging.info("GUI.output_startallocation: " + text)
        elif errorlevel == 2:  # warning
            logging.warning("GUI.output_startallocation: " + text)
        elif errorlevel == 3:  # error
            logging.error("GUI.output_startallocation: " + text)
        elif errorlevel == 4:  # critical
            logging.critical("GUI.output_startallocation: " + text, exc_info=True)
        text_logging.configure(state="normal")
        text_logging.insert("end", "\n" + text)
        text_logging.configure(state="disabled")

    global PATH
    global FILENAME
    global CCALIST
    global CCAALIASES
    global CCATYPE
    global SHEETORDER
    global CHOICES
    global ALOB
    messagebox.showinfo("Info", "Allocation has begun. This make take a while.")
    loginto(1, "Allocation has begun.")
    if PATH == "":
        messagebox.showerror("Error", "Path not inputted.")
        loginto(3, "Path not inputted.")
        return None
    if FILENAME == "":
        messagebox.showerror("Error", "File name not inputted")
        loginto(3, "File name not inputted.")
        return None
    if CCALIST == []:
        messagebox.showerror("Error", "CCAs not inputted.")
        loginto(3, "List of CCAs not inputted.")
        return None
    if SHEETORDER["studentList"] == None:
        messagebox.showerror("Error", "Sheet order of student list not inputted.")
        loginto(3, "Position of list of students in workbook not inputted")
        return None
    if SHEETORDER["CCAList"] == None:
        messagebox.showerror("Error", "Sheet order of CCA list not inputted.")
        loginto(3, "Position of list of CCAs in workbook not inputted")
        return None
    if SHEETORDER["choices"] == None:
        messagebox.showerror("Error", "Sheet order of choices not inputted.")
        loginto(3, "Position of choices of students in workbook not inputted")
        return None
    loginto(1, "Initialising...")
    try:
        ALOB = Allocation(
            path=PATH,
            fileName=FILENAME,
            listOfCCAs=CCALIST,
            CCAAliases=CCAALIASES,
            CCAType=CCATYPE,
            sheetOrder=SHEETORDER,
            numberOfChoices=CHOICES,
        )
    except Exception as e:
        loginto(4, "Error occurred when initialising. Reason: " + str(e))
        return None
    loginto(1, "Finding and opening file...")
    try:
        ALOB.OpenFile()
    except ErroneousParameterError:
        loginto(3, "Could not find file with given path and file name.")
        return None
    except MissingSheetError:
        loginto(3, "No sheets in workbook.")
        return None
    except Exception as e:
        loginto(4, "Could not open file. Reason: " + str(e))
        return None
    loginto(1, "Standardising info...")
    try:
        ALOB.Setup()
    except Exception as e:
        loginto(4, "Unknown error occurred when standardising info. Reason: " + str(e))
        return None
    loginto(1, "Retrieving data...")
    try:
        ALOB.GetData()
    except (MissingSheetError, MissingDataError, ErroneousDataError) as e:
        loginto(3, "Error: Missing sheets/data or erroneous data. Statement: " + str(e))
        return None
    except Exception as e:
        loginto(4, "Unknown error occurred when retrieving data. Reason: " + str(e))
        return None
    loginto(1, "Allocating now...")
    try:
        ALOB.Allocate()
    except Exception as e:
        loginto(4, "Unknown error occurred when allocating students. Reason: " + str(e))
        return None
    try:
        ALOB.SaveToFile()
    except UnableToSaveFileError:
        loginto(3, "Could not save file.")
    except Exception as e:
        loginto(4, "Unknown error when saving file. Reason: " + str(e))
    print("Allocation complete. Check your input workbook file.")
    loginto(1, "Execution complete.")


class Lottery:
    def __init__(self, root):
        def loginto(errorlevel, text):
            if errorlevel == 0:  # debug
                logging.debug("GUI.output_startlottery: " + text)
            elif errorlevel == 1:  # info
                logging.info("GUI.output_startlottery: " + text)
            elif errorlevel == 2:  # warning
                logging.warning("GUI.output_startlottery: " + text)
            elif errorlevel == 3:  # error
                logging.error("GUI.output_startlottery: " + text)
            elif errorlevel == 4:  # critical
                logging.critical("GUI.output_startlottery: " + text, exc_info=True)
            text_logging.configure(state="normal")
            text_logging.insert("end", "\n" + text)
            text_logging.configure(state="disabled")

        loginto(1, "Creating input window.")

        global ALOB  # Get ALOB because all the data is inside there
        if ALOB == None:
            messagebox.showerror("Error", "Students have not been allocated yet.")
            loginto(3, "Allocation not done yet.")
        else:
            if ALOB.Allocated:
                self.entry_inputwindow = tkinter.Toplevel(root)
                self.entry_inputwindow.title("Lottery Options")
                self.label_inputclass = tkinter.Label(
                    self.entry_inputwindow, text="Class:"
                )
                self.label_inputclass.grid(row=1, column=1, pady=5)
                self.entry_inputclass = tkinter.Entry(
                    self.entry_inputwindow, bd=5, width=20
                )
                self.entry_inputclass.grid(row=2, column=1, padx=5, pady=5)
                self.label_inputstudents = tkinter.Label(
                    self.entry_inputwindow, text="Student Indices:"
                )
                self.label_inputstudents.grid(row=3, column=1, pady=5)
                self.entry_inputstudents = tkinter.Entry(
                    self.entry_inputwindow, bd=5, width=20
                )
                self.entry_inputstudents.grid(row=4, column=1, pady=5)
                self.button_getitems = tkinter.Button(
                    self.entry_inputwindow,
                    text="Begin lottery!",
                    command=self.output_getlottery,
                )
                self.button_getitems.grid(row=5, column=1, pady=5)
            else:
                messagebox.showerror("Error", "Students have not been allocated yet.")
                loginto(3, "Allocation not done yet.")

    def output_getlottery(self):
        self._class = self.entry_inputclass.get()
        self._students = self.entry_inputstudents.get()
        print("Students before split:" + str(self._students))
        print("Classes before split:" + str(self._class))
        self.entry_inputwindow.destroy()
        if self._class != "":
            self._class = self._class.split()
            if type(self._class) == str:
                self._class = [self._class]
        else:
            self._class = None
        if self._students != "":
            self._students = self._students.split()
            for student in range(len(self._students)):
                try:
                    self._students[student] = int(self._students[student])
                except ValueError:
                    logging.warning(
                        "GUI.Lottery.output_getlottery: String in self._students"
                    )
        else:
            self._students == None
        print("Students after split:" + str(self._students))
        print("Classes after split:" + str(self._class))
        if self._class == []:
            for student in ALOB.Lottery(_studentIndex=self._students):
                messagebox.showinfo(
                    "Info",
                    "Student Index Number: {0[0]}\nStudent Name: {0[1]}\nCCA: {0[2]}\nShortlist rank in CCA: {0[3]}\nChoice number: {0[4]}\nFitness Score: {0[5]}\nMethod of allocation: {0[6]}".format(
                        student
                    ),
                )
        else:
            lotteryReturn = ALOB.Lottery(
                _studentIndex=self._students, _class=self._class
            )
            print(lotteryReturn)
            if lotteryReturn == None:
                messagebox.showinfo("Info", "No students found.")
            else:
                for _class in range(len(lotteryReturn)):
                    for student in lotteryReturn[_class]:
                        messagebox.showinfo(
                            "Info",
                            "Class: {0}\nStudent Index Number: {1[0]}\nStudent Name: {1[1]}\nCCA: {1[2]}\nShortlist rank in CCA: {1[3]}\nChoice number: {1[4]}\nFitness Score: {1[5]}\nMethod of allocation: {1[6]}".format(
                                self._class[_class], student
                            ),
                        )


def output_startlottery():
    entry_lotterypopup = Lottery(app_frame)
    del entry_lotterypopup


class TemplateWindow:
    def __init__(self, root):
        def loginto(errorlevel, text):
            if errorlevel == 0:  # debug
                logging.debug("GUI.output_createtemplate: " + text)
            elif errorlevel == 1:  # info
                logging.info("GUI.output_createtemplate: " + text)
            elif errorlevel == 2:  # warning
                logging.warning("GUI.output_createtemplate: " + text)
            elif errorlevel == 3:  # error
                logging.error("GUI.output_createtemplate: " + text)
            elif errorlevel == 4:  # critical
                logging.critical("GUI.output_createtemplate: " + text, exc_info=True)
            text_logging.configure(state="normal")
            text_logging.insert("end", "\n" + text)
            text_logging.configure(state="disabled")

        print("Creating window for template.")
        loginto(1, "Creating window")

        global TPOB
        self.entry_inputwindow = tkinter.Toplevel(root)
        self.entry_inputwindow.title("Template Options")
        self.label_winputfilepath = tkinter.Label(self.entry_inputwindow, text="Path:")
        self.label_winputfilepath.grid(row=1, column=1, pady=5)
        self.entry_winputfilepath = tkinter.Entry(
            self.entry_inputwindow, bd=5, width=20
        )
        self.entry_winputfilepath.grid(row=2, column=1, padx=5, pady=5)
        self.label_winputfilename = tkinter.Label(
            self.entry_inputwindow, text="File name:"
        )
        self.label_winputfilename.grid(row=3, column=1, pady=5)
        self.entry_winputfilename = tkinter.Entry(
            self.entry_inputwindow, bd=5, width=20
        )
        self.entry_winputfilename.grid(row=4, column=1, pady=5)
        self.label_inputccas = tkinter.Label(self.entry_inputwindow, text="CCAs:")
        self.label_inputccas.grid(row=5, column=1, pady=5)
        self.entry_inputccas = tkinter.Entry(self.entry_inputwindow, bd=5, width=40)
        self.entry_inputccas.grid(row=6, column=1, padx=5, pady=5)
        self.label_inputsheetorder = tkinter.Label(
            self.entry_inputwindow, text="Admin sheets:"
        )
        self.label_inputsheetorder.grid(row=7, column=1, pady=5)
        self.entry_inputsheetorder = tkinter.Entry(
            self.entry_inputwindow, bd=5, width=20
        )
        self.entry_inputsheetorder.grid(row=8, column=1, pady=5)
        self.label_inputchoices = tkinter.Label(
            self.entry_inputwindow, text="Number of choices:"
        )
        self.label_inputchoices.grid(row=9, column=1, pady=5)
        self.entry_inputchoices = tkinter.Entry(self.entry_inputwindow, bd=5, width=20)
        self.entry_inputchoices.grid(row=10, column=1, pady=5)
        self.button_maketemplate = tkinter.Button(
            self.entry_inputwindow,
            command=self.output_maketemplate,
            text="Create template!",
        )
        self.button_maketemplate.grid(row=11, column=1, pady=5)

    def output_maketemplate(self):
        self.path = self.entry_winputfilepath.get()
        self.filename = self.entry_winputfilename.get()
        self.ccas = self.entry_inputccas.get()
        self.sheetorder = self.entry_inputsheetorder.get()
        self.choices = self.entry_inputchoices.get()
        self.entry_inputwindow.destroy()
        self.ccas = self.ccas.split()
        self.sheetorder = self.sheetorder.split(" ")
        self.choices = self.choices.split(" ")
        self.insheetorder = {
            "studentList": False,
            "healthStats": False,
            "music": False,
            "art": False,
            "special": False,
            "CCAList": False,
            "choices": False,
        }
        for boolean in range(len(self.sheetorder)):
            if self.sheetorder[boolean] == "t":
                self.insheetorder[list(self.insheetorder.keys())[boolean]] = True
        print(self.sheetorder)
        print(self.insheetorder)
        self.inchoices = {"main": 1, "other": None}
        for index in range(len(self.choices)):
            try:
                if index == 0:
                    self.inchoices["main"] = int(self.choices[0])
                elif index == 1:
                    self.inchoices["other"] = int(self.choices[1])
            except ValueError:
                pass
        print(self.choices)
        print(self.inchoices)
        self.template = Template(
            path=self.path,
            fileName=self.filename,
            CCAs=self.ccas,
            sheetOrder=self.insheetorder,
            choices=self.inchoices,
        )
        print("Template created.")
        messagebox.showinfo("Info", "Template created.")


def output_createtemplate():
    entry_templatepopup = TemplateWindow(app_frame)
    del entry_templatepopup


# Setup labels
label_inputfilepath = tkinter.Label(app_frame, text="Path:", anchor="e")

label_inputfilepath.grid(row=1, column=1, padx=10, pady=3)

label_inputfilename = tkinter.Label(app_frame, text="File name:", anchor="e")

label_inputfilename.grid(row=2, column=1, padx=10, pady=3)

label_inputfilewindow = tkinter.Label(
    app_frame, text="Open window to select file:", anchor="e"
)

label_inputfilewindow.grid(row=3, column=1, padx=10, pady=3)

label_inputccalist = tkinter.Label(app_frame, text="List of CCAs:", anchor="e")

label_inputccalist.grid(row=4, column=1, padx=10, pady=2)

label_inputccatype = tkinter.Label(app_frame, text="Type of CCAs:", anchor="e")

label_inputccatype.grid(row=5, column=1, padx=10, pady=2)

label_sheetorderstudentlist = tkinter.Label(
    app_frame, text="Position of student list:", anchor="e"
)

label_sheetorderstudentlist.grid(row=6, column=1, padx=10, pady=2)

label_sheetorderhealthstats = tkinter.Label(
    app_frame, text="Position of health stats:", anchor="e"
)

label_sheetorderhealthstats.grid(row=7, column=1, padx=10, pady=2)

label_sheetordermusic = tkinter.Label(
    app_frame, text="Position of music sheet:", anchor="e"
)

label_sheetordermusic.grid(row=8, column=1, padx=10, pady=2)

label_sheetorderart = tkinter.Label(
    app_frame, text="Position of art sheet:", anchor="e"
)

label_sheetorderart.grid(row=9, column=1, padx=10, pady=2)

label_sheetorderspecial = tkinter.Label(
    app_frame, text="Position of achievements sheet:", anchor="e"
)

label_sheetorderspecial.grid(row=10, column=1, padx=10, pady=2)

label_sheetorderccalist = tkinter.Label(
    app_frame, text="Position of CCA list:", anchor="e"
)

label_sheetorderccalist.grid(row=11, column=1, padx=10, pady=2)

label_sheetorderchoices = tkinter.Label(
    app_frame, text="Position of student choices:", anchor="e"
)

label_sheetorderchoices.grid(row=12, column=1, padx=10, pady=2)

label_choicesmain = tkinter.Label(
    app_frame, text="Number of allowed main choices:", anchor="e"
)

label_choicesmain.grid(row=13, column=1, padx=10, pady=2)

label_choicesother = tkinter.Label(
    app_frame, text="Number of allowed misc choices:", anchor="e"
)

label_choicesother.grid(row=14, column=1, padx=10, pady=2)

# Setup text entry
entry_inputfilepath = tkinter.Entry(bd=5, width=50)

entry_inputfilepath.grid(row=1, column=2, padx=10, pady=3, sticky="w")

entry_inputfilename = tkinter.Entry(bd=5, width=50)

entry_inputfilename.grid(row=2, column=2, padx=10, pady=3, sticky="w")

entry_inputccalist = scrolledtext.ScrolledText(app_frame, width=40, height=5)

entry_inputccalist.grid(row=4, column=2, padx=10, pady=3, sticky="w")

entry_inputccatype = scrolledtext.ScrolledText(app_frame, width=40, height=5)

entry_inputccatype.grid(row=5, column=2, padx=10, pady=3, sticky="w")

entry_sheetorderstudentlist = tkinter.Entry(bd=5, width=10)

entry_sheetorderstudentlist.grid(row=6, column=2, padx=10, pady=3, sticky="w")

entry_sheetorderhealthstats = tkinter.Entry(bd=5, width=10)

entry_sheetorderhealthstats.grid(row=7, column=2, padx=10, pady=3, sticky="w")

entry_sheetordermusic = tkinter.Entry(bd=5, width=10)

entry_sheetordermusic.grid(row=8, column=2, padx=10, pady=3, sticky="w")

entry_sheetorderart = tkinter.Entry(bd=5, width=10)

entry_sheetorderart.grid(row=9, column=2, padx=10, pady=3, sticky="w")

entry_sheetorderspecial = tkinter.Entry(bd=5, width=10)

entry_sheetorderspecial.grid(row=10, column=2, padx=10, pady=3, sticky="w")

entry_sheetorderccalist = tkinter.Entry(bd=5, width=10)

entry_sheetorderccalist.grid(row=11, column=2, padx=10, pady=3, sticky="w")

entry_sheetorderchoices = tkinter.Entry(bd=5, width=10)

entry_sheetorderchoices.grid(row=12, column=2, padx=10, pady=3, sticky="w")

entry_choicesmain = tkinter.Entry(bd=5, width=10)

entry_choicesmain.grid(row=13, column=2, padx=10, pady=3, sticky="w")

entry_choicesother = tkinter.Entry(bd=5, width=10)

entry_choicesother.grid(row=14, column=2, padx=10, pady=3, sticky="w")

# Setup buttons
button_inputfilepath = tkinter.Button(
    app_frame, text="Input", command=output_inputfilepath, padx=10, pady=2
)

button_inputfilepath.grid(row=1, column=3, padx=10, pady=3)

button_inputfilename = tkinter.Button(
    app_frame, text="Input", command=output_inputfilename, padx=10, pady=2
)

button_inputfilename.grid(row=2, column=3, padx=10, pady=3)

button_inputfilewindow = tkinter.Button(
    app_frame, text="Input", command=output_inputfilewindow, padx=10, pady=2
)

button_inputfilewindow.grid(row=3, column=2, padx=10, pady=3)

button_inputccalist = tkinter.Button(
    app_frame, text="Input", command=output_inputccalist, padx=10, pady=2
)

button_inputccalist.grid(row=4, column=3, padx=10, pady=3)

button_inputccatype = tkinter.Button(
    app_frame, text="Input", command=output_inputccatype, padx=10, pady=2
)

button_inputccatype.grid(row=5, column=3, padx=10, pady=3)

button_sheetorderstudentlist = tkinter.Button(
    app_frame, text="Input", command=output_sheetorderstudentlist, padx=10, pady=2
)

button_sheetorderstudentlist.grid(row=6, column=3, padx=10, pady=3)

button_sheetorderhealthstats = tkinter.Button(
    app_frame, text="Input", command=output_sheetorderhealthstats, padx=10, pady=2
)

button_sheetorderhealthstats.grid(row=7, column=3, padx=10, pady=3)

button_sheetordermusic = tkinter.Button(
    app_frame, text="Input", command=output_sheetordermusic, padx=10, pady=2
)

button_sheetordermusic.grid(row=8, column=3, padx=10, pady=3)

button_sheetorderart = tkinter.Button(
    app_frame, text="Input", command=output_sheetorderart, padx=10, pady=2
)

button_sheetorderart.grid(row=9, column=3, padx=10, pady=3)

button_sheetorderspecial = tkinter.Button(
    app_frame, text="Input", command=output_sheetorderspecial, padx=10, pady=2
)

button_sheetorderspecial.grid(row=10, column=3, padx=10, pady=3)

button_sheetorderccalist = tkinter.Button(
    app_frame, text="Input", command=output_sheetorderccalist, padx=10, pady=2
)

button_sheetorderccalist.grid(row=11, column=3, padx=10, pady=3)

button_sheetorderchoices = tkinter.Button(
    app_frame, text="Input", command=output_sheetorderchoices, padx=10, pady=2
)

button_sheetorderchoices.grid(row=12, column=3, padx=10, pady=3)

button_choicesmain = tkinter.Button(
    app_frame, text="Input", command=output_choicesmain, padx=10, pady=2
)

button_choicesmain.grid(row=13, column=3, padx=10, pady=3)

button_choicesother = tkinter.Button(
    app_frame, text="Input", command=output_choicesother, padx=10, pady=2
)

button_choicesother.grid(row=14, column=3, padx=10, pady=3)

button_reset = tkinter.Button(
    app_frame, text="Reset", command=output_resetdata, padx=10, pady=2
)

button_reset.grid(row=15, column=1, padx=10, pady=3)

button_startallocation = tkinter.Button(
    app_frame, text="Start allocation", command=output_startallocation, padx=10, pady=2
)

button_startallocation.grid(row=15, column=2, padx=10, pady=3)

button_lottery = tkinter.Button(
    app_frame, text="Begin lottery.", command=output_startlottery, padx=10, pady=2
)

button_lottery.grid(row=15, column=3, padx=10, pady=3)

button_createtemplate = tkinter.Button(
    app_frame, text="Create template.", command=output_createtemplate, padx=10, pady=2
)

button_createtemplate.grid(row=15, column=4, padx=10, pady=3)

# Keep running
print("Tkinter GUI is running!")
logging.info("GUI: Tkinter GUI is running.")
root.mainloop()

print("Tkinter GUI has stopped.")
logging.info("GUI: Tkinter GUI has stopped.")
