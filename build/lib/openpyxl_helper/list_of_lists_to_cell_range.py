# you could do the same or similar with dictionaries
# entry - puts a list of lists [[1,2],
#                               [3,4]] - a square area of data
# into an excel sheet.
def list_of_lists_to_cell_range(list_of_lists, work_sheet, top_left='A1', bottom_right='unknown'):
    r = 0
    c = 0
    if bottom_right == "unknown":
        bottom_right = number_to_column_letter(len(list_of_lists[0]) + 1) + str(len(list_of_lists) - 1 + int(row_from_coordinate(top_left)))
    for row in work_sheet[top_left:bottom_right]:
        for cell in row:
            print(r)
            print(c)
            print(list_of_lists[r][c])
            cell.value = list_of_lists[r][c]
            c = c + 1
        r = r + 1
        c = 0

def number_to_column_letter(number):
    div = number
    string = ""
    temp = 0
    while div > 0:
        module = (div - 1) % 26
        string = chr(65 + module) + string
        div = int ((div - module) / 26)
    return string

def row_from_coordinate(coordinate):
    number_string = ""
    for character in list(coordinate):
        if represents_integer(character):
            number_string = number_string + character
    return number_string

def represents_integer(integer):
    try:
        int(integer)
        return True
    except ValueError:
        return False
