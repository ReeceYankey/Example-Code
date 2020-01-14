from GradeSheets import *
import pandas as pd
import openpyxl as xl


def UpdateFromCSV(class_names):    
    wb = xl.load_workbook(filename='Galipatia Academic Success Database.xlsx')
    try:
        for name in class_names:
            table = pd.read_csv('{}.csv'.format(name))
            print(type(table['type']))
            ws = wb[name]
            # check type of point system
            if '(WP)' in ws['A2'].value:
                sheet = WeightedSheetHandler(ws)
            else:
                sheet = PointSheetHandler(ws)
            sheet.update(table)
    finally:
        wb.save('updated.xlsx')


if __name__ == '__main__':
    class_names = ['ENGE 1215', 'ENGR 1054', 'CHEM 1035', 'MATH 2204', 'CHEM 1045', 'GEOG 1014']
    UpdateFromCSV(class_names)
