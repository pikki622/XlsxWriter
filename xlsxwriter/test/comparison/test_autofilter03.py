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

        self.set_filename('autofilter03.xlsx')
        self.set_text_file('autofilter_data.txt')

    def test_create_file(self):
        """
        Test the creation of a simple XlsxWriter file with an autofilter.
        This test corresponds to the following examples/autofilter.py example:
        Example 3. Autofilter with a dual filter condition in one of the
        columns.
        """

        workbook = Workbook(self.got_filename)
        worksheet = workbook.add_worksheet()

        # Set the autofilter.
        worksheet.autofilter('A1:D51')

        # Add filter criteria.
        worksheet.filter_column(0, 'x == East or x = South')

        with open(self.txt_filename) as textfile:
            # Read the headers from the first line of the input file.
            headers = textfile.readline().strip("\n").split()

            # Write out the headers.
            worksheet.write_row('A1', headers)

                # Read the rest of the text file and write it to the worksheet.
            for row, line in enumerate(textfile, start=1):

                # Split the input data based on whitespace.
                data = line.strip("\n").split()

                # Convert the number data from the text file.
                for i, item in enumerate(data):
                    try:
                        data[i] = float(item)
                    except ValueError:
                        pass

                # Get some of the field data.
                region = data[0]

                        # Check for rows that match the filter.
                if region not in ('East', 'South'):
                    # We need to hide rows that don't match the filter.
                    worksheet.set_row(row, options={'hidden': True})

                # Write out the row data.
                worksheet.write_row(row, 0, data)

        workbook.close()

        self.assertExcelEqual()
