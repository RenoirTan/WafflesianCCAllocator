# Input Workbook File Format Details

This document is compatible with the Wafflesian CCAllocator Version 1 beta onwards. If you have not read `README.md` yet, do check it out.

## Summary

This document gives basic details on how an input workbook should look like before you run it through the program.

### Administrative sheets

The input workbook must have the following administrative sheets:

1. Student List,
2. CCA List,
3. Choices

Optional administrative sheets include:

4. Health Stats,
5. Music,
6. Art,
7. Special

Each corresponding sheet's name does not have to match with the standard names above. However, you must take note of their position in the workbook, so that the program will know where to retrieve data from.

For example, `Student List` could be called `classlist` in your workbook. However, the program will still know where to retrieve the data of each student from just by telling it where the sheet is. (E.g.: 1)

#### Format of Student List

Student list has a total of 9 columns, namely:

1. Student Number (in consecutive ascending order, otherwise there will be errors)
2. Index Number (same as Student Number)
3. Name
4. ID (another form of identification)
5. Class (not used)
6. Date of Birth (not used)
7. Allocated CCA (used after allocation)
8. Choice # (used after allocation)
9. CCA Rank (used after allocation)

#### Format of CCA List

CCA List has 4 rows:

1. CCA Number
2. CCA Name (can be standard name or nickname)
3. CCA Long Name (for school staffs' identification of CCAs)
4. Allocated (how many students each CCA can take a maximum of)

#### Format of Choices

1. Student Number
2. ID
3. Date of birth (not used)
4. Class
5. Past CCA (not used)
6. Onwards (main choices)
7. Even more onwords (other choices)

#### Format of Health Stats

1. Student Number
2. Name
3. ID
4. Age (not used)
5. Class (not used)
6. Health parameters

During data retrieval, the program will assume that the higher that a student scores at each parameter, the better the student is. However, as weight is an exception to this, `Weight` will be an exception keyword in health parameters.

#### Format of Music & Art

1. ID
2. Class
3. Remarks

#### Format of Special

1. Name
2. Domain
3. CCA

### CCA Shortlist Sheets

Each CCA can have their shortlist sheets. However, they must follow the following format:

1. Rank
2. Name
