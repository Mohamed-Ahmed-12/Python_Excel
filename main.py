from openpyxl import load_workbook
file_name = "ash.xlsx"
workbook = load_workbook(filename=file_name)
sheet = workbook.active
alpha = ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V",
         "W", "X", "Y", "Z"]
number_of_columns = int(input("Enter number of columns you need: "))
print("*** enter title for each columns (in row 1) ****")
for x in alpha:
    if alpha.index(x) != number_of_columns:
        sheet[f"{x}1"] = input(f'{x}1: ')
    else:
        break

number_of_rows = int(input("Enter number of rows you need: "))
first_row = 2 # to start from row 2
for i in range( 2,number_of_rows+2 ):
    for x in alpha:
        if alpha.index(x) != number_of_columns:
            sheet[f"{x}{first_row}"] = input(f"{x}{first_row}: ")
        else:
            break
    print("*"*10)
    first_row = first_row+1

workbook.save(filename="ash.xlsx")