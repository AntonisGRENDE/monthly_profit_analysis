import pandas as pd
import openpyxl
import os

# Load the CSV file
file_path = r""
data = pd.read_csv(file_path)

# Ensure the "Time" column is properly parsed as datetime
data['Time'] = pd.to_datetime(data['Time'], errors='coerce')

# Extract "Month-Year" from the "Time" column
data['Month-Year'] = data['Time'].dt.strftime('%m.%Y')

# Group data by "Month-Year" and calculate Profit and Loss metrics
monthly_summary = data.groupby('Month-Year').agg(
    Total_Profit=('Result', lambda x: x[x > 0].sum()),
    Total_Loss=('Result', lambda x: x[x < 0].sum()),
    Positive_transactions=('Result', lambda x: (x > 0).sum()),
    Negative_transactions=('Result', lambda x: (x < 0).sum()),
    Total_Currency_Conversion_Fees=('Currency conversion fee', 'sum')
).reset_index()

# Ensure currency conversion fees are always positive by taking the absolute value, if needed
monthly_summary['Total_Currency_Conversion_Fees'] = monthly_summary['Total_Currency_Conversion_Fees'].abs()

# Calculate Profit and Loss: Total Profit - Total Loss - Total Currency Conversion Fees
monthly_summary['Profit_and_Loss'] = monthly_summary['Total_Profit'] + monthly_summary['Total_Loss'] - monthly_summary['Total_Currency_Conversion_Fees']

# Reorder columns to place "Profit_and_Loss" second
columns_order = ['Month-Year', 'Profit_and_Loss', 'Total_Profit', 'Total_Loss', 'Positive_transactions', 'Negative_transactions', 'Total_Currency_Conversion_Fees']
monthly_summary = monthly_summary[columns_order]

# Rename columns by replacing underscores with spaces
monthly_summary.columns = [col.replace('_', ' ') for col in monthly_summary.columns]

# Save the monthly summary to a new Excel file
output_file = r"C:\Users\anton\Projects\stonks\monthly_profit_loss_summary.xlsx"
monthly_summary.to_excel(output_file, index=False)

# Open the Excel file using openpyxl to autofit columns and align cells
wb = openpyxl.load_workbook(output_file)
ws = wb.active

# Autofit columns dynamically based on the maximum length of content in each column
for col in ws.columns:
    max_length = 0
    column = col[0].column_letter  # Get the column name
    for cell in col:
        try:
            if len(str(cell.value)) > max_length:
                max_length = len(cell.value)
        except:
            pass
    adjusted_width = (max_length + 2)  # Adding extra space for padding
    ws.column_dimensions[column].width = adjusted_width

# Center-align all cells in the worksheet
for row in ws.iter_rows():
    for cell in row:
        cell.alignment = openpyxl.styles.Alignment(horizontal='center', vertical='center')

# Save the file again after autofitting the columns and aligning the cells
wb.save(output_file)

# Open the Excel file using the default viewer
os.startfile(output_file)

print(f"Monthly Profit and Loss statement saved to {output_file}")