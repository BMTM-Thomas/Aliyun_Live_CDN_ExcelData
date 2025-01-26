import xlwings as xw
import pandas as pd
from xlwings.constants import DeleteShiftDirection

name = "ven319"

try:
    # Read the CSV file into a DataFrame
    df = pd.read_csv(f'{name}.csv')  # Replace 'input.csv' with your CSV file name

    # Convert and save the DataFrame to an Excel file
    df.to_excel(f'{name}.xlsx', index=False)  # Save as 'output.xlsx' (change the file name if needed)

    # Open the workbook
    wb = xw.Book(f'{name}.xlsx')

    # Reference the sheet
    sheet = wb.sheets["Sheet1"]

    # Delete the range of columns
    sheet.range('A:I').api.Delete(DeleteShiftDirection.xlShiftToLeft)
    sheet.range('B:S').api.Delete(DeleteShiftDirection.xlShiftToLeft)
    sheet.range('C:E').api.Delete(DeleteShiftDirection.xlShiftToLeft)
    sheet.range('E:F').api.Delete(DeleteShiftDirection.xlShiftToLeft)
    sheet.range('F:H').api.Delete(DeleteShiftDirection.xlShiftToLeft)
    
    # (f"A2:E1000") = select a range to sort,  Key1=sheet.range("B2").api => sort from which column
    sheet.range(f"A2:E{len(df)+1}").api.Sort(  # +1 due to im not start from header
        Key1=sheet.range("B2").api,
        Order1=1, 
        Orientation=1)
    
    # Read data from the sheet
    data = sheet.range(f'A2:E{len(df)}').value

    next_row = 2
    start_row = 2

    # set cell to 0
    sheet.range(f'L{next_row}').value = 0
    sheet.range(f'M{next_row}').value = 0

    # for loop range from 2 to based on length of excel row
    for i in range(len(df)):

        # if 【分拆项ID】first ID name is similar to the next ID name and 【用量单位】== "GB", then do something
        if sheet.range(f'B{start_row}').value == sheet.range(f'B{start_row + 1}').value and sheet.range(f'D{start_row}').value == "GB" :  
        
            # Write ID
            sheet.range(f'K{next_row}').value = sheet.range(f'B{start_row}').value

            # Sum of Total
            # Total of GB
            sheet.range(f'L{next_row}').value = sheet.range(f'L{next_row}').value + sheet.range(f'C{start_row}').value
            # Total of Prices
            sheet.range(f'M{next_row}').value = sheet.range(f'M{next_row}').value + sheet.range(f'E{start_row}').value
        
        # else if 【分拆项ID】first ID name is similar to the next ID name and 【用量单位】!= "GB", then do something
        elif sheet.range(f'B{start_row}').value == sheet.range(f'B{start_row + 1}').value and sheet.range(f'D{start_row}').value != "GB":
            
            # Set a value 0 instead of "None" value
            if sheet.range(f'L{next_row + 1}').value is None:
                sheet.range(f'L{next_row + 1}').value = 0
                sheet.range(f'M{next_row + 1}').value = 0

            # Write ID
            sheet.range(f'K{next_row + 1}').value = sheet.range(f'B{start_row}').value
            
            # Sum of Total
            # Total of Prices
            sheet.range(f'M{next_row + 1}').value = sheet.range(f'M{next_row + 1}').value + sheet.range(f'E{start_row}').value

        # else 【分拆项ID】first ID name is [NOT similar] to the next ID name
        else:
            
            if sheet.range(f'D{start_row}').value == "GB":
                # Continue Write ID
                sheet.range(f'K{next_row}').value = sheet.range(f'B{start_row}').value
                # Continue Sum of Total
                sheet.range(f'L{next_row}').value = sheet.range(f'L{next_row}').value + sheet.range(f'C{start_row}').value
                sheet.range(f'M{next_row}').value = sheet.range(f'M{next_row}').value + sheet.range(f'E{start_row}').value

            elif sheet.range(f'D{start_row}').value == "万次":

                # Set a value 0 instead of "None" value
                if sheet.range(f'L{next_row + 1}').value is None:
                    sheet.range(f'L{next_row + 1}').value = 0
                    sheet.range(f'M{next_row + 1}').value = 0

                # Continue Write ID
                sheet.range(f'K{next_row + 1}').value = sheet.range(f'B{start_row}').value
                # Continue Sum of Total
                sheet.range(f'M{next_row + 1}').value = sheet.range(f'M{next_row + 1}').value + sheet.range(f'E{start_row}').value


            if sheet.range(f'K{next_row + 1}').value is not None:
                next_row += 2
            else:
                next_row += 1


            # Increment of next_row +1, then set a cell = 0, instead of "None" value
            if i != len(df)-1:
                sheet.range(f'L{next_row}').value = 0
                sheet.range(f'M{next_row}').value = 0

        start_row += 1

    # Save the workbook (if needed)
    wb.save()

    # Close the workbook
    wb.close()

except Exception as e:
    print(e)
    wb.save()
    wb.close()
