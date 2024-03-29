﻿using Aspose.Cells;
using System;
using System.Drawing;

namespace ExcelComparator
{
    class Program
    {
        static void Main(string[] args)
        {
            if (args.Length < 2)
            {
                Console.WriteLine("Please provide filenames for both Excel files.");
                return;
            }

            string file1Path = args[0];
            string file2Path = args[1];

            string licPath = "Aspose.Total.lic";
            Aspose.Cells.License lic = new Aspose.Cells.License();
            lic.SetLicense(licPath);

            Workbook workbook1 = new Workbook(file1Path);
            Workbook workbook2 = new Workbook(file2Path);

            // Sort the worksheets by the first column
            SortSheetByFirstColumn(workbook1.Worksheets[0]);
            SortSheetByFirstColumn(workbook2.Worksheets[0]);

            // Highlight missing columns in each worksheet compared to the other
            HighlightMissingColumns(workbook1.Worksheets[0], workbook2.Worksheets[0]);
            HighlightMissingColumns(workbook2.Worksheets[0], workbook1.Worksheets[0]);

            // Compare the two workbooks
            CompareWorkbooks(workbook1, workbook2, false);
            CompareWorkbooks(workbook2, workbook1, true);

            // Save the updated workbooks with highlighted differences
            workbook1.Save(file1Path);
            workbook2.Save(file2Path);
        }

        static void SortSheetByFirstColumn(Worksheet worksheet)
        {
            // Get the range of cells to be sorted (excluding header)
            int startRow = 1; // Assuming the header is in the first row
            int startColumn = 0; // First column
            int endRow = worksheet.Cells.MaxDataRow + 1; // Add 1 to include the last row
            int endColumn = worksheet.Cells.MaxDataColumn + 1; // Add 1 to include the last column

            // Create a range based on the specified cells
            Range rangeToSort = worksheet.Cells.CreateRange(startRow, startColumn, endRow, endColumn);

            DataSorter sorter = worksheet.Workbook.DataSorter;
            sorter.Key1 = 0;
            sorter.Sort(worksheet.Cells, startRow, startColumn, endRow, endColumn);
        }

        static void HighlightMissingColumns(Worksheet sourceWorksheet, Worksheet targetWorksheet)
        {
            for (int columnIndex = 0; columnIndex <= sourceWorksheet.Cells.MaxDataColumn; columnIndex++)
            {
                string columnName = sourceWorksheet.Cells[0, columnIndex].StringValue;
                if (!WorksheetContainsColumn(targetWorksheet, columnName))
                {
                    Style styleMissing = targetWorksheet.Cells.GetCellStyle(0, columnIndex);
                    styleMissing.ForegroundColor = System.Drawing.Color.Red;
                    styleMissing.Pattern = BackgroundType.Solid;
                    StyleFlag flag = new StyleFlag();
                    //targetWorksheet.Cells.Columns[columnIndex].ApplyStyle(styleMissing, new StyleFlag { FontColor = true });
                    sourceWorksheet.Cells[0, columnIndex].SetStyle(styleMissing);
                }
            }
        }

        static bool WorksheetContainsColumn(Worksheet worksheet, string columnName)
        {
            for (int i = 0; i <= worksheet.Cells.MaxDataColumn; i++)
            {
                if (worksheet.Cells[0, i].StringValue.Equals(columnName))
                {
                    return true;
                }
            }
            return false;
        }

        static void CompareWorkbooks(Workbook workbook1, Workbook workbook2, bool OnlyHighlightAdditionalRows)
        {
            Worksheet worksheet1 = workbook1.Worksheets[0];
            Worksheet worksheet2 = workbook2.Worksheets[0];

            Style styleMismatch = workbook1.CreateStyle();
            styleMismatch.ForegroundColor = Color.Yellow;
            styleMismatch.Pattern = BackgroundType.Solid;
            StyleFlag flag = new StyleFlag();

            // Iterate through each row in worksheet1
            for (int rowIndex = 1; rowIndex <= worksheet1.Cells.MaxDataRow; rowIndex++) // Start from the second row to skip the header
            {
                string key = worksheet1.Cells[rowIndex, 0].StringValue; // Get the value from the first column as key

                // Find the row in worksheet2 with the same key
                int targetRowIndex = FindRowIndex(worksheet2, key);

                if (targetRowIndex == -1)
                {
                    // Highlight missing row in worksheet2
                    //worksheet1.Cells.Rows[rowIndex].ApplyStyle(styleMismatch, flag);
                    // worksheet1.Cells[rowIndex, worksheet1.Cells.MaxDataColumn].SetStyle(styleMismatch);
                    //Aspose.Cells.Range range = worksheet1.Cells.CreateRange(rowIndex, 1, 1, worksheet1.Cells.MaxDataColumn);
                    //range.SetStyle(styleMismatch);
                    Style styleMissing = worksheet1.Cells.GetCellStyle(rowIndex, 0);
                    styleMissing.ForegroundColor = System.Drawing.Color.Red;
                    styleMissing.Pattern = BackgroundType.Solid;
                    //targetWorksheet.Cells.Columns[columnIndex].ApplyStyle(styleMissing, new StyleFlag { FontColor = true });
                    worksheet1.Cells[rowIndex, 0].SetStyle(styleMissing);
                }
                else
                {
                    if (OnlyHighlightAdditionalRows)
                    {
                        continue;
                    }
                    // Compare values of the rows
                    for (int columnIndex = 1; columnIndex <= worksheet1.Cells.MaxDataColumn; columnIndex++)
                    {
                        string value1 = worksheet1.Cells[rowIndex, columnIndex].StringValue;
                        string columnName = worksheet1.Cells[0, columnIndex].StringValue;
                        int targetColumnIndex = FindColumnIndex(worksheet2, columnName);

                        if (targetColumnIndex != -1)
                        {
                            string value2 = worksheet2.Cells[targetRowIndex, targetColumnIndex].StringValue;

                            if (!value1.Equals(value2))
                            {
                                // Highlight the cell in yellow in both worksheets
                                worksheet1.Cells[rowIndex, columnIndex].SetStyle(styleMismatch);
                                worksheet2.Cells[targetRowIndex, targetColumnIndex].SetStyle(styleMismatch);
                            }
                        }
                    }
                }
            }
        }

        static int FindRowIndex(Worksheet worksheet, string key)
        {
            // Iterate through each row in the worksheet
            for (int rowIndex = 1; rowIndex <= worksheet.Cells.MaxDataRow; rowIndex++) // Start from the second row to skip the header
            {
                // Compare the value in the first column (assuming it's the key)
                if (worksheet.Cells[rowIndex, 0].StringValue.Equals(key))
                {
                    return rowIndex; // Return the index if found
                }
            }
            return -1; // Return -1 if not found
        }

        static int FindColumnIndex(Worksheet worksheet, string columnName)
        {
            // Iterate through each column in the header row
            for (int columnIndex = 0; columnIndex <= worksheet.Cells.MaxDataColumn; columnIndex++)
            {
                // Compare the column name
                if (worksheet.Cells[0, columnIndex].StringValue.Equals(columnName))
                {
                    return columnIndex; // Return the index if found
                }
            }
            return -1; // Return -1 if not found
        }
    }
}
