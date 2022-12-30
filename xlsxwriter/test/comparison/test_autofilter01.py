###############################################################################
#
# Tests for XlsxWriter.
#
# SPDX-License-Identifier: BSD-2-Clause
# Copyright (c), 2013-2022, John McNamara, jmcnamara@cpan.org
#

from ..excel_comparison_test import ExcelComparisonTest
from ...workbook import Workbook


class TestCompareXLSXFiles(ExcelComparisonTest):
    """
    Test file created by XlsxWriter against a file created by Excel.

    """

    def setUp(self):

        self.set_filename('autofilter01.xlsx')
        self.set_text_file('autofilter_data.txt')

    def test_create_file(self):
        """
        Test the creation of a simple XlsxWriter file with an autofilter.
        This test corresponds to the following examples/autofilter.py example:
        Example 1. Autofilter without conditions.
        """

        workbook = Workbook(self.got_filename)
        worksheet = workbook.add_worksheet()

        # Set the autofilter.
        worksheet.autofilter('A1:D51')

        with open(self.txt_filename) as textfile:
                # Read the text file and write it to the worksheet.
            for row, line in enumerate(textfile):

                # Split the input data based on whitespace.
                data = line.strip("\n").split()

                # Convert the number data from the text file.
                for i, item in enumerate(data):
                    try:
                        data[i] = float(item)
                    except ValueError:
                        pass

                # Write out the row data.
                worksheet.write_row(row, 0, data)

        workbook.close()

        self.assertExcelEqual()
