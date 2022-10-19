#######.PY PROGRAM FOR APPENDING COORDS TO WORK EXCEL SHEET
##AUTHOR >> SB
import time

import openpyxl  # for ExcEl WOrKB00k
from colorama import Fore, Back, Style  # for console colors
import time

# get the start time
st = time.time()

m_list = []  # to hold the map coordinates

# Open the sample.txt file containing the map coords obtained from manual data entry
with open('files/sample.txt') as file:
    # contents = file.read()
    # print(contents)

    # Read the contents of the text file
    for address in file.readlines():
        # print(address.strip())  # the strip method removes the blank line as each line has a newline character(\n)

        # before appending the address lets split it at the comma and space
        new_address = address.strip().split(', ')

        # let's convert them to type float
        first_no = float(new_address[0])
        second_no = float(new_address[1])

        # round of both numbers
        first_no = round(first_no, 6)
        second_no = round(second_no, 6)

        # concatenate both values and ", " in between
        save_address = str(first_no) + ", " + str(second_no)

        # append the new short address to the list
        m_list.append(save_address)

file.close()  # close the file document after use

# load the workbook
workbook = openpyxl.load_workbook("files/sample.xlsx")

sheet = workbook.active  # obtain the active sheet in the workbook

# to hold the cells that the user wants to skip and leave blank
blank_cells_list = []

print("WELCOME TO >> " + Fore.LIGHTBLUE_EX + "AUTO_FILLER" + "4 ExCeL \n" + Style.RESET_ALL)
# Request for input of blank cells
while True:
    blank_cell = input(
        Style.BRIGHT + Fore.GREEN + "Enter cell to leave blank \n" + "0 and 1 will be ignored \n" + Fore.RED + "Exit by pressing enter at the next prompt >> " + Style.RESET_ALL)

    # break loop if no value is entered
    if blank_cell == "":
        break
    else:
        # append blank_cells value to blank_cells list
        blank_cells_list.append(blank_cell)

# Print the entered blank cells list
print(Back.BLUE + "Blank cells to leave: ", blank_cells_list)


# function to ad the cell values to the active sheet in the workbook
def add_values(cells):
    # Appending the LIST to the WorkBooK
    i = 2  # to start at the second row as the first row is used for column titles

    # to insert a blank value at the location of the requested blank cells list
    for y in cells:
        m_list.insert(int(y) - 2, "")

    # print(Fore.CYAN + "<< The new list >> \n" + Style.RESET_ALL, m_list)

    # now append the reformatted list to the workbook
    for a_row in m_list:
        cell = sheet.cell(row=i, column=1)
        cell.value = a_row

        i += 1  # increment to enter next value at new row


add_values(blank_cells_list)  # call the add_values() to add values to the cell

# get the end time
et = time.time()

print(">>>>>>>>>DONE<<<<<<<<<<")
print("Estimated Time Taken = " + str(round(et - st, 2)) + "seconds")

# save the workbook with the new values
workbook.save(filename="files/sample.xlsx")
