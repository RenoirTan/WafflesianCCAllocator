# The Wafflesian CCAllocator Version 1.4.0 beta


## Introduction

The Wafflesian CCAllocator is a program that will make allocating CCAs to students fast and simple. With just your input excel workbook and a few clicks, you can allocate CCAs to students! All in mere seconds.

Depending on the number of students, CCAs and wide range of parameters you can input, allocation can be as fast as 2.5 seconds!

## Getting started

This program's executable `.exe` works on `Windows` only, unfortunately. To run this on other operating systems, please check out the section below.

### Prerequisites and Installation

The executable is only available for Windows. To get the program, simply download it and click on it.

If you are using another operating system or want to use the python source code, please read the instructions below.

1. Download the zip file of the source code.
2. Extract all files in the zip file.
3. Download a version of Python that is 3.4 or above that is compatible with your OS/distribution. (Download the MSI installer to make installation simpler.)
4. Install all optional features. (The `Python` program will be using it later.)
5. __(optional)__ Create a virtual environment by navigating to the desired folder where you want the virtual environment to be with the terminal and enter:

   ```
   pip install virtualenv
   virtualenv (directory) <- input with desired name of virtual environment
   source (directory)/bin/activate <- activate virtual environment (enter this line every time you need to use the program)
   ```

6. Install all required modules in `/CCA/scripts/requirements.txt`

   To do this, open up your terminal and navigate to the folder (`.../CCA/scripts/`) where the `Python` program is. Type in:
   `pip install -r requirements.txt`

   Alternatively, you could enter into the terminal:

      ```
      pip install logging
      pip install openpyxl
      pip install pathlib
      pip install traceback
      ```

### Setting up the input workbook

To ensure compatibility with the program, make sure that the input file is in line with the format in outlined in `FILEFORMAT.md`. Alternatively, you can follow the example workbook in `.../xlsxfiles/template` for the python program or `.../template/` for the executable program.

### Running the program (GUI)

If you wish to run the GUI program, run `CCAllocatorApp.exe` or `/CCA/scripts/CCAllocatorApp.py` and then, simply enter in the parameters and click `Start Allocation` at the bottom. Below is a list of what each input means:

1. Path __(Compulsory)__: Path to where the input file is. E.g.: `C:/ExamplePath/FolderInTheFolder/CCAAllocations`
2. File name __(Compulsory)__: Name of the input file. E.g.: `students_and_ccas.xlsx`. __Make sure that your input file is a Microsoft Excel Workbook which ends with `.xlsx`.__
Alternatively, you could use the graphic file selector labelled `Open window to select file:`.
3. List of CCAs __(Compulsory)__: List of CCAs and their nicknames as they appear in the input file (CCAs separated by line breaks, nicknames separated by spaces). E.g.:
```
CCA nickname1 nickname2
GuitarEnsemble guitar_cca
Painting
Swimming swim swimmingcca
Programming Infocomm
```
4. Type of CCAs: List of CCAs with their type. List of allowed types: `m (music), a (art), s (sport), b (basic)` E.g.:
```
CCA type
GuitarEnsemble m
Painting a
Swimming s
Programming b
```
5. Position of *sheets*: Position of administrative sheets in the input file.

   List of administrative sheets:
   1. Student List __(Compulsory)__: Sheet of list of students and basic info about them.
   2. Health Stats: Sheet of list of each students health data.
   3. Music: Sheet of list of students who show potential or interest in music CCAs.
   4. Art: Sheet of list of students who show potential or interest in art CCAs.
   5. Achievements: Sheet of each student's CCA in which they have achieved the best results or achievements in.
   6. CCA list __(Compulsory)__: Sheet of list of CCAs and basic information about them.
   7. Choices __(Compulsory)__: Sheet of list of students' choices.

6. Number of *choices*:

   1. Main choices: Number of main CCAs each student is allowed to choose.
   2. Other choices: Same as above but for miscellaneous choices.

Once you have completed inputting all necessary information, press `Start Allocation` below.

However, if you want to use `CCAllocator.py` as a module with a separate `Python` file to input parameters and run the algorithm, you can check out `example.py` in `.../CCA/scripts/`.

__Once allocation has begun and finished, you can check out your input workbook. The allocated CCAs would appear in `Student List`.__ If you wish to, you can now press the `Begin Lottery` button, input the parameters and then have a popup surprise for your students.

### Using the python module

If you want better control over the program, you can use python to allocate students. To do this, import `/CCA/scripts/CCAllocator.py` over and start coding.

In total, there are 7 parameters and 7 functions/methods. The program uses Object-Oriented Programming and requires you to initiate an instance of the `Allocation` class to get started.

#### Parameters (Variables)

When initialising an instance for allocation, you can enter the following *kwargs*:

1. path __(string, Compulsory)__: Path to where the input workbook is.
2. fileName __(string, Compulsory)__: File name. Must include suffix (.xlsx)
3. listOfCCAs __(list, Compulsory)__: List of CCAs in their standard name.
4. CCAAliases __(dictionary, Optional)__: Dictionary of list of each CCA's nicknames. If the CCA does not have a nickname, no need to include it in the dictionary. Keys: Standard name of CCA, Value: [List of CCA's nicknames]
5. CCAType __(dictionary, Optional)__: CCA Type. Keys: Standard name of CCA, Value: Type of CCA, can be "m", "a", "s" or "b" (music, art, sports, basic) respectively.
6. sheetOrder __(dictionary, Compulsory)__: Where all the admin sheets are located. This tells the program where the important sheets are located. Example:
   ```
   {
       "studentList":1, # Sheet with list of students
       "healthStats":2, # Sheet with health of students
       "music":3, # Sheet of students with music interest
       "art":4, # Sheet of students with art interest
       "special":5, # Sheet of students with exceptional abilities in the CCA or has been in the CCA before
       "CCAList":6, # Sheet of CCAs with info about them
       "choices":7, # Sheet with the choices of the students
   }
   ```
   You can call CCAllocator.SHEETORDER to get the format.
7. numberOfChoices __(dictionary, Compulsory)__: Dictionary of how main and other choices students get. Example:
   ```
   {
      "main":9, # 9 main choices
      "other":2 # 2 optional choices
   }
   ```
   You can call CCAllocator.CHOICES to get the format.

#### Methods

Once you have initialised the instance with the above parameters, you can use the following methods in that order:

1. `Allocation.OpenFile()`: Opens file
2. `Allocation.Setup()`: Sets up attributes with information in the file
3. `Allocation.GetData()`: Obtains and configures data
4. `Allocation.Allocate()`: Allocates students
5. `Allocation.SaveToFile()`: Saves data to input workbook
6. `Allocation.Lottery(_print=bool, _studentIndex=int or list of ints, _class=string or list of strings)`: Retrieves data of students after allocation.

### Creating a template

If you to create a template workbook, you can do so by pressing `Create Template` in the GUI or entering in `template = CCAllocator.Template(kwargs)`.

Creating a new template in the GUI:

There are 5 inputs:

1. Path: Path to where the template should be.
2. File name: Desired name of template. Must end in .xlsx.
3. CCAs: List of CCAs. Separate with spaces.
4. Sheet Order: Which admin sheets will appear in the template. You configure which sheets appear by using *t* as True and *f* as False, denoting whether the sheet will appear in the following order:

   1. studentList
   2. healthStats
   3. music
   4. art
   5. special
   6. CCAList
   7. choices

   For example, if *t f f f f t t*, only *studentList*, *CCAList* and *choices* will appear.
5. Choices: How many choices each students is entitled. The numbers for main choices and other choices are separated by a space. For example: *9 2* means that each student can have 9 main choices and 2 other choices.

Keyword arguments for `CCAllocator.Template()`:

1. path __(string)__: String to where the template should be.
2. fileName __(string)__: Desired file name of template. Must end in .xlsx
3. CCAs __(list)__: List of CCAs.
4. sheetOrder __(dictionary)__: Which admin sheets will appear in the template. It is the same as sheetOrder in Allocation(kwargs) but all values are booleans.
5. choices __(dictionary)__: How many main choices and other choices each student is entitled to. Same as choices in Allocation(kwargs):

#### Example

If you need an example, you can open `/CCA/scripts/example.py` and read it to understand how to use the module.

For examples on creating a template, you can check out `/CCA/scripts/create_template.py`

### Built with

[Python](https://www.python.org) - *Language*

[Tkinter](https://effbot.org/tkinterbook/) - *Effbot documentation*

[Black](https://github.com/psf/black) - *Code formatter*

### Contributing

Please read `CONTRIBUTING.md` for details on how to contribute.

### Versioning

This project uses `SemVer` versioning. Check out `CHANGELOG.md` for changelogs.

### Authors and contribution

__Renoir Tan__ - *Wrote all the code lol* - [DerperorWaffle](https://github.com/DerperorWaffle)

__Tian Xiang Cheng__ - *Contributed ideas as to how the students should be allocated*

### License

We aren't qualified for licensing please don't steal.

### Acknowledgements

__Mrs Neo__ for her invaluable guidance and advice :)
