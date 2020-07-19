import csv
import xlrd
import xlsxwriter
import os
import timeit as timer

# Took out Quantity Unit
# Index: 2
TITLELIST = ['Quantity number', 'Packaging', 'Brand', 'Categories', 'Energy in kcals',
             'Grams of fat', 'Grams of saturated fat', 'Grams of carbohydrates', 'Grams of sugar',
             'Grams of salt', 'Nutrition score – France', 'Nutri-score letter grade',
             'Serving size quantity', 'Serving size unit']

# Indexes of the information requested in order of request
INDEXLIST = [9, 10, 12, 16, 71, 74, 75, 110, 111, 127, 172, 54, 41, 40]

# Longest barcode + 2
LENGTHCONST = 59 + 2


def sortFunc(bc):
    length = len(bc)

    bc_num = bc;

    # Removes leading zeros so string can be converted into int successfully
    while bc_num[0] == '0':
        bc_num = bc_num[1:]

    # Sets every barcode to an equal length of 61 (unless it had leading zeros in which case it ll be proportionally
    # smaller)
    order_num = int(bc_num) * 10 ** (LENGTHCONST - length)

    # adds the original length of the barcode back to te number in order to create a unique and accurate number that
    # can be used or ordering EX: barcode 1 should go before barcode 100 because it's shorter (these are the barcode
    # rules I've observed in the .csv file) 1 * 10^61 = 100 * 10^58 || 1 * 10^61 + 1 < 100 * 10^58 + 3
    order_num = order_num + length
    return order_num


def getBarcodes(loc):
    wb = xlrd.open_workbook(loc)
    sheet = wb.sheet_by_index(0)
    return_array = []
    for i in range(1, sheet.nrows):
        val = sheet.cell_value(i, 0)
        if val != '':
            return_array.append(val)
    return_array.sort(key=sortFunc)
    return return_array


def read_csv(barcodes=[], window=None):
    stored_array = []
    return_array = []
    for bc in barcodes:
        return_array.append([])

    search_index = 0

    while search_index < len(barcodes):
        with open('data/en.openfoodfacts.org.products.csv', 'r', encoding='utf8', errors='ignore') as csv_file:
            csv_reader = csv.reader(csv_file, delimiter='\t')

            i = 0

            next(csv_reader)
            for line in csv_reader:
                i += 1

                if window.window_exited:
                    return None

                progress_val = int((float(i)/float(1389841))*100)
                progress_val = min(progress_val, 100)
                window.progress_bar.setValue(progress_val)
                # stored_array.append(line)
                if str(line[0]) == barcodes[search_index]:
                    print("Barcode: " + str(line[0]) + " || Product Name: " + line[7] + "\r")
                    for index in INDEXLIST:
                        return_array[search_index].append(line[index])
                    search_index = search_index + 1
                    if search_index >= len(barcodes):
                        window.reset_window()
                        return return_array

        window.throw_error('Barcode Error', 'Please check this barcode, it is not in the database: ' + str(barcodes[search_index]))
        search_index = search_index + 1
    window.reset_window()
    return return_array


def write_xl(da, path=os.path.expanduser('~/Documents/output_workbook.xlsx')):
    out_wb = xlsxwriter.Workbook(path)
    out_sheet = out_wb.add_worksheet()

    out_sheet.set_column('A:' + chr(64 + len(TITLELIST)), 20)

    bold = out_wb.add_format({'bold': True})

    i = 0
    for title in TITLELIST:
        out_sheet.write(chr(65 + i) + '1', TITLELIST[i], bold)
        i = i + 1

    i = 0
    for data_set in da:
        j = 0
        for data in data_set:
            out_sheet.write(chr(65 + j) + str(i + 2), data)
            j = j + 1
        i = i + 1
    out_wb.close()
