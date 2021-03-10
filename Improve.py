import xlsxwriter
file_name = input("Enter File name: ex: file.xlsx : ")
pattern = input("Enter Pattern: ex: 01H : ")
start_number = int(input('Start Number : ex: 4654900 : '))
end_number = int(input('End Number : ex: 4656900 : '))
first_row = input("Enter First Row name : ex: IGT8_0.6M / LED : ")
second_row = input("Enter Second Row name : ex : 9W / 3W : ")
third_row = input("Enter Third Row name : ex: 6500K / 865 / 830 : ")
fourth_row = input("Enter Fourth Row name : ex: E27 / B22 : ")

workbook = xlsxwriter.Workbook(file_name)
worksheet = workbook.add_worksheet()
# number assign as quantity
number = end_number
j = 0
for i in range(start_number, number+1):
    my_list = (f"{pattern}{i}")
    j = (i - start_number)
    worksheet.write(j, 0, j)
    worksheet.write(j, 1, first_row)
    worksheet.write(j, 2, second_row)
    worksheet.write(j, 3, third_row)
    worksheet.write(j, 4, fourth_row)
    worksheet.write(j, 5, "1901")
    worksheet.write(j, 6, my_list)

    print(my_list)

workbook.close()