import pandas as pd
import os
from openpyxl import load_workbook
import numpy as np
from openpyxl import Workbook
from copy import copy


# xls = pd.ExcelFile('ا. - لحن اللى التوزيع')
# df1 = pd.read_excel(xls, 'ك مارجرحس حداثق المعادى')
# df2 = pd.read_excel(xls, 'ك مارجرجس بكوتسيكا')

# file_name = "ا. - لحن نى اثنوس تيرو.xlsx"
# xls = pd.read_excel(file_name, sheet_name = "ك مارمرقس بالمعادى")
# xls_1 = pd.read_excel(file_name, sheet_name = "ك الانبا بيشوى بزهراء المعادى")
# print (xls)

# xls.to_excel("output.xlsx", engine='xlsxwriter', mode='a', sheet_name="")  
# xls_1.to_excel("output_1.xlsx")  

# main_file = "ا. - لحن نى اثنوس تيرو.xlsx"

# book = load_workbook(main_file)
# writer = pd.ExcelWriter(main_file, engine='openpyxl') 
# writer.book = book
# sheet_1 = book.worksheets[0]
# print(sheet_1['B10'].value)

# for row in sheet_1.iter_rows(min_row=6, max_col=6, max_row=40, values_only=True):
#     listo = list(row)

#     print (listo[1])
#     print (listo[2])
#     print (listo[3])
#     print (listo[4])
#     print (listo[5])
#     input("e")
        

# print (sheet_1.cell(row=10, column=2))
# dicy = {}
# for ws in book.worksheets:
#     print (ws.title)
#     x = input("do u want this?")
#     if x == 1:
#         dicy[ws.title] = ws
        

# writer.sheets = dicy

# writer.sheets = dict( (ws.title, ws) for ws in book.worksheets)

# data_filtered.to_excel(writer, "Main", cols=['Diff1', 'Diff2'])
# sh1 =  list(writer.sheets.keys())
# print (sh1[0])
# sh1 = writer.sheets[sh1[0]]
# print (sh1)
# print (sh1.cell(2,3))

# print ("ev")
# writer.save()

# book.save("ev.xlsx")

def reverse(string):
    # strings =(list(string))
    out = ""
    n = len(string)
    for i in range (len(string)):
        out += string[n-1-i]
    return out

def fill_random(file_name):
    '''
        Fill data randomly for validation purposes
    '''
    book = load_workbook(file_name)
    writer = pd.ExcelWriter(file_name, engine='openpyxl') 
    writer.book = book

    for sheet in  book.worksheets:
        print (f'Working with file: {reverse(file_name)} \tsheet: {reverse(sheet.title)}')
        if sheet['A5'].value != "م":
            print ("Error: the sixth row is not the start!!")
        cell_index = 6
        for row in sheet.iter_rows(min_row=6, max_col=6, values_only=True):
            listo = list(row)
            # print (listo)
            #exit when reach the end
            if listo[1] == None:
                break
            
            if listo[0] == None and listo[1] != None:
                print("\nDOUBLE!")
                print (f'WARNING: file name: {reverse(file_name)} \tsheet: {reverse(sheet.title)} \tCell: {cell_index} \tFound a DOUBLE!')
                print (f'name: {reverse(listo[1])}\n\n')
                cell_index += 1
                continue

            pos = 'C'+str(cell_index)
            if sheet[pos].value == None:
            # if  sheet[pos]._style.fillId == 0 :
                rand_int = np.random.randint(0,4)
                sheet['C'+str(cell_index)] = rand_int

                rand_int = np.random.randint(0,3)
                sheet['D'+str(cell_index)] = rand_int

                rand_int = np.random.randint(0,3)
                sheet['E'+str(cell_index)] = rand_int

            cell_index+=1
    writer.save()


def check_values_bound(file_name=None):
    '''
        For each sheet we have, check the bounded values
        if they exceed (4,3,3), then print out the file name/sheet name/ and the row
    '''
    book = load_workbook(file_name)
    writer = pd.ExcelWriter(file_name, engine='openpyxl') 
    writer.book = book

    for sheet in book.worksheets:
        print (f'Working with sheet: {(sheet.title)}')
        if sheet['A5'].value != "م":
            print ("Error: the sixth row is not the start!!")

        for row in sheet.iter_rows(min_row=6, max_col=6, values_only=True):
            listo = list(row)
            print (listo)
            #exit when reach the end
            if listo[0] == None:
                break

            four, three_1, three_2 = listo[2:5]
            # listo[5] = int(four) + int(three_1) + int(three_2)  
            if int(four) > 4 or int(three_1) > 3 or int(three_2) >3:
                print (f'HUGE ERROR: file name: {file_name} \tsheet: {sheet.title} \trow_index: {listo[0]} Validation error!')
                input("Ev")
    
    # cells = book.worksheets[0]['A1':'A40']
    # for cell in cells:
        
    #     print (cell[0].style_id)
    #     print (cell[0]._style.fillId)


def merge_separated (source_1, source_2):
    wb = Workbook()
    
    book = load_workbook(source_1)
    writer = pd.ExcelWriter(source_1, engine='openpyxl') 
    writer.book = book

    dicy = {}
    for sheet in book.worksheets:
        # dicy[sheet.title] = sheet
        sh = wb.create_sheet(sheet.title)
        # for row_src in sheet.iter_rows(min_row=1, max_col=6, values_only=True):
        for row_src in sheet.rows:
            # listo = list(row_src)
            # if listo[1] == None and listo[0] == None and listo[2] == None and listo[3] == None and listo[4] == None:
            #     break

            for cell in row_src:
                if cell == None:
                    continue
                try:
                    new_cell = sh.cell(row=cell.row, column=cell.col_idx,
                            value= cell.value)
                    if cell.has_style:
                        new_cell.font = copy(cell.font)
                        new_cell.border = copy(cell.border)
                        new_cell.fill = copy(cell.fill)
                        new_cell.number_format = copy(cell.number_format)
                        new_cell.protection = copy(cell.protection)
                        new_cell.alignment = copy(cell.alignment)
                except AttributeError:
                    print ('sorry')
        destination = source_1.split('.')[1] + " + " +source_2.split('.')[1] + ".xlsx"

        wb.save(destination)
        return


        # max_rows = 1
        # for row_src in sheet.iter_rows(min_row=6, max_col=6, values_only=True):
        #     listo = list(row_src)
        #     if listo[1] == None and  listo[0] == None :
        #         break
        #     max_rows += 1

        # for i in range (1,max_rows+7):
        #     sh['A'+str(i)] = sheet['A'+str(i)].value
        #     sh['B'+str(i)] = sheet['B'+str(i)].value
        #     sh['C'+str(i)] = sheet['C'+str(i)].value
        #     sh['D'+str(i)] = sheet['D'+str(i)].value
        #     sh['E'+str(i)] = sheet['E'+str(i)].value
        #     sh['F'+str(i)] = sheet['F'+str(i)].value



    book = load_workbook(source_2)
    writer = pd.ExcelWriter(source_2, engine='openpyxl') 
    writer.book = book

    for sheet in  book.worksheets:
        # dicy[sheet.title] = sheet
        sh = wb.create_sheet(sheet.title)
        max_rows = 1
        for row_src in sheet.iter_rows(min_row=6, max_col=6, values_only=True):
            listo = list(row_src)
            if listo[1] == None and  listo[0] == None :
                break
            max_rows += 1

        for i in range (1,max_rows+1+7):
            sh['A'+str(i)] = sheet['A'+str(i)].value            
            sh['B'+str(i)] = sheet['B'+str(i)].value
            sh['C'+str(i)] = sheet['C'+str(i)].value
            sh['D'+str(i)] = sheet['D'+str(i)].value
            sh['E'+str(i)] = sheet['E'+str(i)].value
            sh['F'+str(i)] = sheet['F'+str(i)].value

    destination = source_1.split('.')[1] + " + " +source_2.split('.')[1] + ".xlsx"
    
    wb.save(destination)
    return destination

def print_info (file_name, sheet_name):
    print (f'Working with file: \t{file_name} \tin sheet: \t{sheet_name}.')

def fill_main(main_file, source_1, mian_extra = True):
    print (f'Destination file: \t{main_file}')
    print (f'Source file: \t\t{source_1}')

    src_book = load_workbook(source_1)
    src_writer = pd.ExcelWriter(source_1, engine='openpyxl') 
    src_writer.book = src_book

    dst_book = load_workbook(main_file)
    dst_writer = pd.ExcelWriter(main_file, engine='openpyxl') 
    dst_writer.book = dst_book
    dst_sheets = list(dst_book.sheetnames)
    dst_sheets_dic = {}
    for src_sheet in src_book.worksheets:
        print_info (source_1, src_sheet.title)

        # find the same sheet in the main file!
        # if it does not exist, print an error message!
        if src_sheet.title in dst_sheets:
            skip_5_rows = 5
            dst_sheet = dst_book[src_sheet.title]
            for row_src in src_sheet.rows:
                if skip_5_rows > 0:
                    skip_5_rows -= 1
                    continue
                for cell in row_src:
                    if cell == None:
                        continue
                    try:
                        if mian_extra:
                            new_cell = dst_sheet.cell(row=cell.row, column=cell.col_idx, value= cell.value)
                            
                        else:
                            if cell.col_idx == 1 or cell.col_idx == 2:
                                continue
                            new_cell = dst_sheet.cell(row=cell.row, column=cell.col_idx+4, value= cell.value)
                            
                        if cell.has_style:
                            new_cell.font = copy(cell.font)
                            new_cell.border = copy(cell.border)
                            new_cell.fill = copy(cell.fill)
                            new_cell.number_format = copy(cell.number_format)
                            new_cell.protection = copy(cell.protection)
                            new_cell.alignment = copy(cell.alignment)
                    except AttributeError:
                        # print ('sorry allignment error occured!')
                        print("\nDOUBLE!")
                        print (f'WARNING: Source File: {reverse(source_1)} \tsheet: {reverse(src_sheet.title)} \tRow: {cell.row} \tFound a DOUBLE!\n\n')
                        break
            dst_sheets_dic [dst_sheet.title] = dst_sheet
        else:
            print (f'Error: sheet \t{src_sheet.title} \tdoes not exist in file: \t{main_file}')
            input ("Evram!!")
            dst_writer.sheets = dst_sheets_dic
            dst_writer.save()
            return

    dst_writer.sheets = dst_sheets_dic
    dst_writer.save()

    

    

if __name__ == "__main__":
    
    if os.listdir("dst/") == []:
        print (f'Error: no destination files!')

    if os.listdir("src_main/") == []:
        print (f'Error: no mian source files!')
    
    if os.listdir("src_extra/") == []:
        print (f'Error: no extra source files!')
        
    dst_path = "dst/" + os.listdir("dst/")[0]
    for src_folders in ["src_main/", "src_extra/"]:
        srcs = os.listdir(src_folders)
        if src_folders.split("_")[1] == "main/":
            mian_extra = True
        else:
            mian_extra = False
        for src in srcs:
            src_path = src_folders + src
            fill_main(dst_path, src_path, mian_extra=mian_extra)
    

    # dest = "استمارة تقييم مهرجان 2020.xlsx"
    # source_1 = "ا.ابونا اغسطينوس كامل - لحن تين او اوشت (( او )) طاى شورى.xlsx"
    # source_2 = "ا.ابرام عادل - لحن تين او اوشت (( او )) طاى شورى.xlsx"
    # source_3 = "ا.بيشوى عادل - لحن اللى التوزيع.xlsx"

    # for file_path in os.listdir("fill_these/"):
    #     main_file = "fill_these/" + file_path
    #     fill_random(main_file)
    # check_values_bound(main_file)
    # merged = merge_separated (source_1, source_2)

    # fill_main(dest, source_1)

    