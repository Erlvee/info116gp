import xlsxwriter
import re

def file_handler(filename):
    file_name = open("C:/Users/erlend/Documents/Obligatorisk arbeid/INFO116/txtDataDistricts/" + filename, "r")
    if file_name.mode == 'r':
        read_file = file_name.read().splitlines() #splits into list based on \n lines
        list_names = [x.strip(' ') for x in read_file] #remove spaces in list
        names_string = ''.join(list_names)
    return names_string

def protege_list(string):
    noNum_string = ''.join(i for i in string if not i.isdigit()) #remove numbers from {stringmethod}
    mod_string = noNum_string.replace("-", "_").replace(".", "_").replace(" ", "_").replace("(", "").replace("__", "_") #modify string {stringmethod}
    splitstring = mod_string.split(")") #Split into list
    list_of_names = list(map(lambda i: i[ : -1], splitstring)) #remove last character due to it being "_"
    return list_of_names

def write_excel(file, worksheet, row, col):
    row = row
    col = col
    for items in file:
        worksheet.write(row, col, items)
        row += 1

def main():
    converted_txtfile = file_handler("countries_vm.txt")
    pro_list = protege_list(converted_txtfile)

    workbook = xlsxwriter.Workbook('C:/Users/erlend/Documents/Obligatorisk arbeid/INFO116/project/vinmonopolet_countries.xlsx')
    worksheet = workbook.add_worksheet()
    write_excel(pro_list, worksheet, 0, 1)
    workbook.close()
    

if __name__ == "__main__":
    main()







