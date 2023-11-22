import csv
import shutil
import openpyxl


# Replace 'source.xlsx' and 'destination_copy.xlsx' with the actual file paths
source_file = 'marksheet.xlsx'

# Replace 'your_file.csv' with the actual path to your CSV file
file_path = 'student-names.csv'

# Replace 'your_file.xlsx', 'Sheet1', 'A', 1, and 'New Value' with your actual file path, sheet name, 
# column letter, row number, and the value you want to add.
sheet_name = 'Sheet1'
row_number = 5


def run():
    process_csv(file_path)


def add_value_to_excel_cell(source_file, sheet_name, row_number, column_letter, value):
    try:
        # Load the workbook
        workbook = openpyxl.load_workbook(source_file)

        # Select the specified sheet
        sheet = workbook[sheet_name]

        # Update the cell with the new value
        sheet[column_letter + str(row_number)] = value

        # Save the changes
        workbook.save(source_file)

        print(f'Successfully added {value} to cell {column_letter}{row_number} in {source_file}')
    except FileNotFoundError:
        print(f"File not found: {source_file}")
    except Exception as e:
        print(f"An error occurred: {e}")


# duplicate file
def duplicate_excel_file(source_file, destination_file):
    try:
        # Copy the Excel file
        shutil.copyfile(source_file, destination_file)
        print(f'Successfully duplicated {source_file} to {destination_file}')
    except FileNotFoundError:
        print(f"Source file not found: {source_file}")
    except PermissionError:
        print(f"Permission error. Check if you have the necessary permissions.")
    except Exception as e:
        print(f"An error occurred: {e}")


# generate filenames from csv
def process_csv(file_path):
    try:
        with open(file_path, 'r') as csv_file:
            # Create a CSV reader object
            csv_reader = csv.reader(csv_file)
            
            # Read the header row
            header = next(csv_reader)
            print(f'CSV Header: {header}')
            
            # Loop through the rows
            for row in csv_reader:
                print(row)  # You can process each row here as needed
                # add to excel
                add_value_to_excel_cell(source_file, sheet_name, row_number, 'B', row[2])
                add_value_to_excel_cell(source_file, sheet_name, row_number, 'D', row[0])
                # duplicate file
                destination_file = "marksheets/feedback_" + row[0] + ".xlsx"
                duplicate_excel_file(source_file, destination_file)

    except FileNotFoundError:
        print(f"File not found: {file_path}")
    except Exception as e:
        print(f"An error occurred: {e}")




if __name__ == "__main__": 
    run()


