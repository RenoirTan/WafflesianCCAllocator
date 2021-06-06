from CCAllocator import *
import os
import pathlib

path = pathlib.Path(os.getcwd())
path = pathlib.Path(*path.parts[: len(path.parts) - 1])
path = pathlib.Path(path, "xlsxfiles")
path = str(path)
filename = "example_template.xlsx"
CCAs = ["CCA1", "CCA2"]
sheetorder = {
    "studentList": True,
    "healthStats": True,
    "music": True,
    "art": True,
    "special": True,
    "CCAList": True,
    "choices": True,
}
choices = {"main": 9, "other": 2}

template = Template(
    path=path, fileName=filename, CCAs=CCAs, sheetOrder=sheetorder, choices=choices
)
