# VBA-challenge
Data Visualization and Analysis Bootcamp - Challenge 2

## Files
The VBA script is contained in the file VBA_challenge.vba. The four PNG files are screenshots of the results of running the script in the provided workbook Multiple_year_stock_data.xlsx (one for each worksheet in the workbook). 

## Note on source
Lines 3 through 7 of the VBA script contain the following code:

    Dim ws As Worksheet
    
        For Each ws In ThisWorkbook.Worksheets

            ws.Select

The idea to use a For loop over the set of all worksheets in the workbook (in order to run the macro on all worksheets simultaenously), as well as the specific syntax in these lines, are adapted from code in the answers here: https://www.mrexcel.com/board/threads/is-there-a-way-to-apply-a-macro-across-all-worksheets.997398/
