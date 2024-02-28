from athlete_data import AthleteData
from athlete_data import fillExcelWorksheet
from copy import copy
from openpyxl import Workbook, load_workbook
from selenium.webdriver.common.by import By


# Returns the index of the option with the given string
def findOptionIndex(elem, name: str):
    option_index = -1
    options = elem.find_elements(By.TAG_NAME, "option")
    numOptions = len(options)
    for index in range(0, numOptions):
        option = options[index]
        if name in option.text:
            option_index = index
            break
    
    return option_index



# Builds and returns the array of points for elite athletes
def buildElitePointsArray():
    elite_points = [100, 92, 86, 82, 80]
    for i in reversed(range(1, 80)):
        elite_points.append(i)

    print(elite_points)

    return elite_points

# Copies the cells from source_sheet to target_sheet
def copy_cells(source_sheet, target_sheet):
    for (row, col), source_cell in source_sheet._cells.items():
        target_cell = target_sheet.cell(column=col, row=row)

        target_cell._value = source_cell._value
        target_cell.data_type = source_cell.data_type

        if source_cell.has_style:
            target_cell.font = copy(source_cell.font)
            target_cell.border = copy(source_cell.border)
            target_cell.fill = copy(source_cell.fill)
            target_cell.number_format = copy(source_cell.number_format)
            target_cell.protection = copy(source_cell.protection)
            target_cell.alignment = copy(source_cell.alignment)

        if source_cell.hyperlink:
            target_cell._hyperlink = copy(source_cell.hyperlink)

        if source_cell.comment:
            target_cell.comment = copy(source_cell.comment)

    target_sheet.conditional_formatting = source_sheet.conditional_formatting


# Sorting function for athletes
def AthleteSorting(athlete: AthleteData):
    return athlete.timeSecs
