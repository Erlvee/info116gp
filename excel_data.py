import xlsxwriter
import re

#string = "France_subdistrict()Gevrey-Chambertin (205)Nuits-St.-Georges (105)Bourgogne (88)Volnay (88)ChambolleMusigny (83)Pommard (79)Morey Saint-Denis (58)Aloxe-Corton (50)Øvrige (40)Morgon (38)Beaune Premier Cru (34)Vougeot (32)Beaujolais-Villages (31)Côte de Nuits (31)Vosne-Romanée (31)Beaujolais (30)Savigny-les-Beaune (28)Fleurie (25)Santenay (25)Moulin-à-Vent (18)Beaune (15)Monthélie (15)Charmes-Chambertin (13)Marsannay (13)Mercurey (13)Santenay 1.er Cru (13)Chénas (12)Chambertin (11)Chassagne-Montrachet (11)Côte de Brouilly (11)Côte de Nuits-Villages (10)Flagey-Echézeaux (10)Hautes-Côtes de Nuit (10)Côte de Beaune (9)Régnié (9)Hautes-Côtes de Beaune (8)Juliénas (7)Rully (7)SaintAubin (7)Brouilly (6)Mercurey Premier Cru (6)Auxey-Duresses (5)Fixin (5)Givry (5)Maranges (5)Chiroubles (4)Côte Chalonnaise (4)Maranges Premier Cru (4)Saint-Amour (4)Saint-Romain (4)Chorey les Beaune (3)Mâcon (3)Pernand-Vergelesses (3)Mazis-Chambertin (2)Meursault (2)Meursault-Blagny (2)Bourgogne Épineuil (1)Côtes d'Auxerre (1)Ladoix (1)Musigny (1)Mâcon-Cruzille (1)Mâcon Verzé (1)Puligny-Montrachet (1)"
file1= open('C:/Users/erlend/Documents/Obligatorisk arbeid/GitHub/INFO116/project/txtDataDistricts/districts_france.txt',"r")
file2= open('C:/Users/erlend/Documents/Obligatorisk arbeid/GitHub/INFO116/project/txtDataDistricts/districts_france.txt',"r")
file3= open('C:/Users/erlend/Documents/Obligatorisk arbeid/GitHub/INFO116/project/txtDataDistricts/subdistrict_france.txt',"r")
file4= open('C:/Users/erlend/Documents/Obligatorisk arbeid/GitHub/INFO116/project/txtDataDistricts/countries.txt',"r")


def file_convert(file):
    if file.mode == 'r':
        read_file = file.read().splitlines()
        list2 = [x.strip(' ') for x in read_file]
        dnames_string = ''.join(list2)
    return dnames_string

def protege_list(string):
    mod_string = string.replace("-", "_").replace(".", "_").replace(" ", "_").replace("(", "")
    splitstring = mod_string.split(")")
    list_of_names = list(map(lambda i: i[ : -1], splitstring)) 
    return list_of_names

def del_last(string):
    remove_numbers = ''.join(i for i in string if not i.isdigit())
    return remove_numbers

def make_xls(list, column, worksheet):
    row = 0
    col = column

    for items in list_names:
        worksheet.write(row, col, items)
        row += 1
    return worksheet

def write_excel(file, worksheet, row, col):
    row = row
    col = col
    for items in file:
        worksheet.write(row, col, items)
        row += 1



def main():
    converted_file = file_convert(file1)
    converted_file3 = file_convert(file3)
    converted_countries = file_convert(file4)
    c_countries = re.sub(r"(\w)([A-Z])", r"\1 \2", converted_countries).split(' ')
    print(c_countries)


    districts_france = del_last(converted_file)
    d_france = protege_list(districts_france)
    print(d_france)

    subdis_france = del_last(converted_file3)
    subd_france = protege_list(subdis_france)
    print(subd_france)


    workbook = xlsxwriter.Workbook('C:/Users/erlend/Documents/Obligatorisk arbeid/GitHub/INFO116/project/vinmonopolet_data.xlsx')
    worksheet = workbook.add_worksheet()

    write_excel(d_france, worksheet, 0, 0)
    write_excel(subd_france, worksheet, 0, 3)
    write_excel(c_countries, worksheet, 0, 6)
    
    

    workbook.close()
    

if __name__ == "__main__":
    main()







