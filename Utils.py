
def calculate_ascii(col, col_count, use_third, first_letter, third_letter):
    """
    Calculates the ascii-betical letter code corresponding to a certain index. Only applies in cases with fewer
    than or equal to three letters.
    :param col: the 
    :param col_count: 
    :param use_third: 
    :param first_letter: 
    :param third_letter: 
    :return: 
    """
    if col <= 26:
        # if it's under 26 columns, just use a single letter
        ascii_col = chr(col + 64)
    elif use_third:
        if col_count > 26:
            # first_letter describes the coordinate of what the first letter should be -
            # every 26 iterations, it increases by one to switch the first letter up by one
            first_letter += 1
            # col_count keeps track of what column you're at in the current first_letter iteration
            col_count = 1
        if first_letter > 90:
            third_letter += 1
            first_letter = 65
        ascii_col = chr(third_letter) + chr(first_letter) + chr((col_count + 64))

        col_count += 1
    else:
        # if it's over 26 columns, you have to calculate two different letters
        if col_count > 26:
            # first_letter describes the coordinate of what the first letter should be -
            # every 26 iterations, it increases by one to switch the first letter up by one
            first_letter += 1
            # col_count keeps track of what column you're at in the current first_letter iteration
            col_count = 1

        ascii_col = chr(first_letter) + chr((col_count + 64))

        if ascii_col == 'ZZ':
            use_third = True

        col_count += 1
    return ascii_col, col_count, use_third, first_letter, third_letter


def create_remove_sheets(workbook, sheet_names):
    """
    Creates specific sheets in a workbook, and deletes the prior ones if they already exist.
    :param workbook: an Excel workbook to be modified
    :param sheet_names: an array of strings specifying the names of the sheets to be created
    """
    for sheet in sheet_names:
        if sheet in workbook.sheetnames:
            remove_sheet = workbook.get_sheet_by_name(sheet)
            workbook.remove_sheet(remove_sheet)

        workbook.create_sheet(sheet)


def find_index_column(sheet, name, num):
    """
    Finds the column by which patients should be separated in a document (for instance, finding the column that 
    specifies the MRN).
    :param workbook: an Excel workbook to be searched
    :param name: a string specifying the name of the column that is to be found
    :param num: the row in which the value will be found
    :return: the letter value of the column
    """
    for idx in range(1, 26):
        if sheet[chr(idx + 64) + str(num)].value == name:
            index_col = chr(64 + idx)
            break
    return index_col
