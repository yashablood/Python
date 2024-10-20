from openpyxl import load_workbook

def load_data(file_path):
    # Load the workbook and the specific sheet
    workbook = load_workbook(file_path)
    sheet = workbook["Data"]

    # Print out the labels and dates to verify the sheet is loaded correctly
    print("Loading data from 'Data' sheet...")
    
    # Print labels in column B and dates in row 1
    labels = [sheet[f'B{i}'].value for i in range(2, 19)]
    dates = [sheet.cell(row=1, column=j).value for j in range(3, sheet.max_column+1)]
    
    print("Labels:", labels)
    print("Dates:", dates)

def update_sheet(file_path, data, selected_date):
    # Load the workbook and sheet
    workbook = load_workbook(file_path)
    sheet = workbook["Data"]

    # Find the column corresponding to the selected date
    date_column = None
    for col in range(3, sheet.max_column + 1):  # Starting from column C (index 3)
        if sheet.cell(row=1, column=col).value == selected_date:
            date_column = col
            break

    if date_column is None:
        print(f"Date '{selected_date}' not found in 'Data' sheet")
        return
    
    # Map the input fields from the UI to the corresponding rows in the "Data" sheet
    sheet.cell(row=2, column=date_column).value = data.get("Days without Incident", "N/A")
    sheet.cell(row=3, column=date_column).value = data.get("Haz ID's", "N/A")
    sheet.cell(row=4, column=date_column).value = data.get("Safety Gemba Walk", "N/A")
    sheet.cell(row=5, column=date_column).value = data.get("7S (Zone 26)", "N/A")
    sheet.cell(row=6, column=date_column).value = data.get("7S (Zone 51)", "N/A")
    sheet.cell(row=7, column=date_column).value = data.get("Errors", "N/A")
    sheet.cell(row=8, column=date_column).value = data.get("PCD Returns", "N/A")
    sheet.cell(row=9, column=date_column).value = data.get("Jobs on Hold", "N/A")
    sheet.cell(row=10, column=date_column).value = data.get("Productivity", "N/A")
    sheet.cell(row=11, column=date_column).value = data.get("OTIF %", "N/A")
    sheet.cell(row=12, column=date_column).value = data.get("Huddles", "N/A")
    sheet.cell(row=13, column=date_column).value = data.get("Truck Fill %", "N/A")
    sheet.cell(row=14, column=date_column).value = data.get("Recognitions", "N/A")
    sheet.cell(row=15, column=date_column).value = data.get("MC Compliance", "N/A")
    sheet.cell(row=16, column=date_column).value = data.get("Cost Savings", "N/A")
    sheet.cell(row=17, column=date_column).value = data.get("Rever's", "N/A")
    sheet.cell(row=18, column=date_column).value = data.get("Project's", "N/A")

    # Save the workbook after updating
    workbook.save(file_path)
    print(f"Data for date '{selected_date}' updated successfully.")
