import os
import unittest

import poi4py
class ExcelTest(unittest.TestCase):

    def test_read(self):
        poi4py.start_jvm()
        file_path = os.path.dirname(__file__) + '/test.xlsx'
        file_path_out = os.path.dirname(__file__) + '/test_out.xlsx'
        print(file_path)
        workbook = poi4py.open_workbook(file_path)
        sheet = workbook.getSheetAt(0)
        for i in range(3, sheet.getLastRowNum() + 1):
            row = sheet.getRow(i)

            for i in range(0, row.lastCellNum):
                cell = row.getCell(i)
                if cell:
                    print('fire')
                    cell.setCellFormula('1+1')
        poi4py.save_workbook(file_path_out, workbook)
        poi4py.shutdown_jvm()
