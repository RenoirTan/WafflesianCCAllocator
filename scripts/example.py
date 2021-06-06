"""
example.py
==========
Test python file to test out CCAllocator.py
"""

import os  # Get our current path
import pathlib  # Module we will be using to modify paths. You can use any method but we recommend pathlib.
import traceback  # In case something goes wrong. Tell me lol

import CCAllocator as ca  # CCAllocator module. As it is in the same folder, no directory changing would be needed

PATH = pathlib.Path(os.getcwd())  # Get path. This will return us with './CCA/scripts'
PATH = pathlib.Path(*PATH.parts[: len(PATH.parts) - 1])  # Remove '/scripts'
PATH = pathlib.Path(PATH, "xlsxfiles")  # Add 'xlsxfiles'
PATH = str(
    PATH
)  # PATH must be a string. You can use any path editing modules but PATH must end up as a string.

FILENAME = "unallocated.xlsx"  # File name of example workbook

# List of CCAs' standard names
CCALIST = [
    "BAD",
    "BAS",
    "FEN",
    "CC",
    "CRI",
    "FLO",
    "GOLF",
    "HOC",
    "JUD",
    "POLO",
    "RUG",
    "SAIL",
    "SHO",
    "SOF",
    "SQU",
    "SWI",
    "TAB",
    "TEN",
    "TNF",
    "BB",
    "NPCC",
    "NCC",
    "RC",
    "O1",
    "O2",
    "SE",
    "RV",
    "GE",
    "RIMB",
    "RICO",
    "RP",
    "RD",
]

# Dictionary of CCAs and their list of aliases (nicknames)
CCAALIASES = {
    "O1": ["01SCOUT"],  # E.g. O1 could also be called 01SCOUT in the workbook
    "O2": ["02SCOUT"],
    "RICO": ["CO"],
    "RD": ["Debater", "raffles_debate"],
}

# CCAs' type. Each CCA can have only one type (for now)
# 'm' stands for music
# 'a' for art
# 's' for sport
# 'b' for basic. Writing for basic CCAs would be unnecessary as the program will assume that the CCA is basic if no category was given in the first place.
CCATYPE = {
    "SE": "m",
    "RV": "m",
    "GE": "m",
    "RIMB": "m",
    "RICO": "m",
    "BAD": "s",
    "BAS": "s",
    "FEN": "s",
    "CC": "s",
    "CRI": "s",
    "FLO": "s",
    "GOLF": "s",
    "HOC": "s",
    "JUD": "s",
    "POLO": "s",
    "RUG": "s",
    "SAIL": "s",
    "SHO": "s",
    "SOF": "s",
    "SQU": "s",
    "SWI": "s",
    "TAB": "s",
    "TEN": "s",
    "TNF": "s",
}

# Position of admin sheets in the workbook
# If a sheet doesn't exist, there is no need to add it in and add None, program will assume so.
# If you don't know how SHEETORDER looks like, just call in ca.SHEETORDER.
SHEETORDER = {
    "studentList": 1,  # Compulsory, exists in position 1
    "healthStats": 2,  # Not compulsory
    "music": 3,  # NC
    "art": None,  # NC
    "special": 4,  # NC
    "CCAList": 5,  # Compulsory
    "choices": 6,  # Compulsory
}

# How many choices each student can have
# You can use ca.CHOICES as well
CHOICES = {"main": 9, "other": 2}

try:
    # Create object for allocation
    allocation_object = ca.Allocation(
        path=PATH,
        fileName=FILENAME,
        listOfCCAs=CCALIST,
        CCAAliases=CCAALIASES,
        CCAType=CCATYPE,
        sheetOrder=SHEETORDER,
        numberOfChoices=CHOICES,
    )
    allocation_object.OpenFile()  # Open the file
    allocation_object.Setup()  # Setup variables
    allocation_object.GetData()  # Obtain data
    allocation_object.Allocate()  # Allocate CCAs
    allocation_object.SaveToFile()  # Save to file
    allocation_object.Lottery(
        _print=True, _class=["1J", "1B"], _studentIndex=[1, 4, 9, 16, 25]
    )
    # Lottery, collect all students info and print out. Check help(ca.Allocation.Lottery) for more info
except:
    traceback.print_exc()  # In case any boo boo surfaces we can check
finally:
    x = input(
        "Continue>>> "
    )  # End button because python.exe exits immediately after program is done or errors occur. This lets us have a chance to see everything going on in the program.
