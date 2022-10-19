'''
for a_row in m_list:  # Iterate through m_list as we append each value from the list into its respective row

#    i += 1  # to increment to move to the next row
#   print("The value of i >> ", i)

    print("The value of skipped value", skipped_value)
    if skipped_value == "":
        print(Back.GREEN + "No skipped value" + Style.RESET_ALL)
    else:
        print(Back.RED + "Skipped value exists" + Style.RESET_ALL)
        cell = sheet.cell(row=i, column=1)
        cell.value = skipped_value

        skipped_value = ""  # reset

        i += 1

    # using the naive method to check if value exists in the list
    if str(i) in cells:
        print(Fore.BLUE + ">> The cell number exists <<" + Style.RESET_ALL)
        skipped_value = a_row
        continue
    else:
        cell = sheet.cell(row=i, column=1)
        cell.value = a_row

    # to add

'''