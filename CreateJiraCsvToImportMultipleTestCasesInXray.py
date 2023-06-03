import codecs
import csv
import logging
import os
import sys
import time

import pandas as pd
from openpyxl import load_workbook


class CreateJiraCsvToImportMultipleTestCasesInXray:
    """
    This class converts the xlsx files (test cases from ALM) and saves them as csv file in the output folder.
    """

    def __init__(self, input_folder, output_folder):
        self.input_folder = input_folder
        self.output_folder = output_folder

    def convert(self):

        for filename in os.listdir(self.input_folder):

            if filename.endswith(".xlsx"):
                input_path = os.path.join(input_folder, filename)
                output_path = os.path.join(output_folder, filename[:-5] + ".csv")
                df = pd.read_excel(input_path)
                df.to_csv(output_path, index=False)

        print("Converting xlsx files from [input folder] and save in [output folder] as csv - successfully ")

class XlsxSetTestcaseNumber:
    """
    This class creates a new csv file --> combined.csv
    This csv file contains all test cases (the test cases are appended one below the other).
    Additionally, the issue ID for the test cases is generated correctly (The issue ID = test case number).
    Issue ID 7 = Test case 7 (Must be present in each step of the test case)
    """

    def __init__(self, file_list):
        self.file_list = file_list
        self.numer_of_files = len(file_list)
        self.counter_for_files = 1

    def process_xlsx_files(self):
        for file_name in self.file_list:
            file_path = os.path.join('input', file_name)

            # XLSX-Datei mit openpyxl öffnen
            wb = load_workbook(file_path)
            sheet = wb.active

            # Nach Header-Spalte "Step Name" suchen
            header_col_index = None
            for col_idx, cell in enumerate(sheet[1]):
                if cell.value == "Step Name":
                    header_col_index = col_idx + 1  # Spaltenindex mit 1-basiertem Index

            if header_col_index is not None:
                # In der Spalte A ab der zweiten Zeile iterieren
                for row in sheet.iter_rows(min_row=2, min_col=header_col_index, max_col=header_col_index):
                    cell_value = row[0].value  # Wert der Zelle in Spalte A

                    # "1" in die Zelle schreiben und fortsetzen, bis eine leere Zelle erreicht wird
                    if cell_value is not None:
                        row[0].value = self.counter_for_files

                    else:
                        break

            self.counter_for_files = self.counter_for_files + 1

            # XLSX-Datei speichern und schließen
            wb.save(file_path)
            wb.close()

        return self.file_list


class ReadXlsxToCSV:
    def __init__(self, csv_folder, output_folder, output_csv_name, xlsx_filenames):

        self.csv_folder = csv_folder
        self.output_csv_name = output_csv_name
        self.output_folder = output_folder
        self.xlsx_filenames = xlsx_filenames
        self.all_rows = []

    def process_csv_files(self):
        for filename in os.listdir(self.csv_folder):
            if filename.endswith('.xlsx'):
                file_path = os.path.join(self.csv_folder, filename)
                time.sleep(0.2)
                print(" - XLSX file found :   ", filename)

                wb = load_workbook(file_path)
                sheet = wb.active

                # Step Name | Summary | Test Type | Expected Result

                sheet.insert_cols(2)
                sheet.cell(row=1, column=2).value = "Summary"
                sheet.cell(row=2, column=2).value = filename

                sheet.insert_cols(3)
                sheet.cell(row=1, column=3).value = "Test Type"
                sheet.cell(row=2, column=3).value = "Manual"

                for i, row in enumerate(sheet.iter_rows(values_only=True)):
                    modified_row = list(row)
                    self.all_rows.append(modified_row)

            wb.save(self.output_csv_name)
            wb.close()

        output_file = os.path.join(self.output_folder, self.output_csv_name)

        with open(output_file, 'w', newline='') as csvfile:
            writer = csv.writer(csvfile, delimiter=';')
            writer.writerows(self.all_rows)

        print('The combined CSV file has been created: {}' + output_file)


class CSVReader:
    """
    In this class the collected csv files (test cases) will be
    appended to each other and a new file will be created - this includes the header of every testcase!!
    The class has a method that iterates through the first column of the csv and searches for value [first_column_key].
    The return value of this method is a list, this list contains the numbers of the row the word
    [first_column_key]. This list contains ALL row numbers the next class will take this list manipulate it and delete
    the proper lines.
    """

    def __init__(self, filename, first_column_key):
        self.filepath = filename
        self.lines_to_delete = []
        self.first_column_key = first_column_key

    def read_csv(self):
        i = -1
        with open(self.filepath, 'r') as csvfile:
            reader = csv.reader(csvfile, delimiter=";")
            for row_num, row in enumerate(reader, start=1):
                if row:
                    i = i + 1

                if row[0] == "Step Name":
                    print(" - Column A 'Step name' found in row [{i}] of the CSV-File.")
                    self.lines_to_delete.append(i)

        print("Found the header in fallowing lines :",self.lines_to_delete)
        return self.lines_to_delete


class SubtractListElements:
    """
    This class is passed a list.
    This is manipulated so that it can be used to remove the corresponding lines from the csv.
    REASON WHY: If a line is deleted form csv , for example the line 6 then the other lines slip after.
    The first element stays as it is, the 2 element slips 1 after the 3 element slips 2 after and so on
    From the list [4, 13, 19, 25, 48, 64, 71, 84, 99, 107]
    Becomes the list [4, 12, 17, 22, 44, 59, 65, 77, 91, 98]
    With this list we can delete the proper lines ( Headline form each attached Testcases
    --> Because we only need the headline on time in the fist line.
    """

    def __init__(self, value_list):
        self.value_list = value_list

    def subtract_elements(self):
        result_list = [self.value_list[0]]  # Das erste Element bleibt unverändert

        for i in range(1, len(self.value_list)):
            subtracted_value = self.value_list[i] - i
            result_list.append(subtracted_value)

        print("Manipulated List with the proper header numbers to delete : ", result_list )
        return result_list


class CSVDeleteLines:
    """
    In this class the csv file is passed to remove the corresponding rows.
    The class creates a new csv whose columns are JIRA compliant.
    """

    def __init__(self, file_path):
        self.file_path = file_path
        self.rows = []

    def read_csv(self):
        with open(self.file_path, 'r', newline='') as csvfile:
            reader = csv.reader(csvfile, delimiter=";")
            self.rows = list(reader)

    def delete_rows(self, rows_to_delete):
        for row_index in rows_to_delete:
            if row_index < len(self.rows):
                del self.rows[row_index]

    def save_csv(self, output_file_path):

        try:
            with open(output_file_path, 'w', newline='') as csvfile:
                writer = csv.writer(csvfile, delimiter=";")
                writer.writerows(self.rows)

        except:
            sys.exit("ERROR: The file multiple_testcases_file_for_jira.csv is already open.\n"
                     "Close the file and run the programm again.\n"
                     "NOTE: The file will be overwritten - to prevent overriding store the file in a other location.\n"
                     )
        print("Headlines deleted to have only one headline in teh file multiple_testcases_file_for_jira.csv")
class ChangeCsvCodingFromAsiToUTF8:
    """
    In this class the csv file is passed, which was adapted for JIRa.
    In order for all umlauts to be displayed correctly in JIRA, the file must conform to the UTF-8 standard.
    The class changes the encoding of the file so that a JIRA import does not show <?> in the text.
    """

    def __init__(self, csv_name):
        self.csv_name = csv_name

    def convert_encoding(self):
        # Actual coding
        aktuelle_codierung = 'cp1252'  # Example: ANSI (Windows-1252)

        with codecs.open(self.csv_name, 'r', encoding=aktuelle_codierung) as datei:
            inhalt = datei.read()

        # Open file and save it as UTF-8-coded file
        with codecs.open(self.csv_name, 'w', encoding='utf-8') as datei:
            datei.write(inhalt)

        print(f'File: {csv_name} was saved in UTF-8 format')


if __name__ == '__main__':

    folders = os.listdir()
    if "output" in folders:
        pass
    else:
        sys.exit("ERROR: The [output] folder does not exist in the current directory.\n"
                 "Please create a folder in the current directory and start the program again.\n")

    count_xlsx_files = 0

    if "input" in folders:

        files_in_input_folder = os.listdir('input/')

        for file in files_in_input_folder:
            if file.endswith(".xlsx"):
                count_xlsx_files += 1
        if count_xlsx_files > 0:
            pass
        else:
            sys.exit("ERROR: Der Ordner [input] muss mindestens ein XLSX-File beinhalten (Testfall).\n "
                     "Bitte fügen sie mindestens einen testfall in das Verzeichnis.\n")

    else:
        sys.exit("ERROR: The [input] folder must contain at least one XLSX file (test case).\n"
                 "Please create a folder in the current directory and start the program again.\n"
                 "NOTE: The folder must contain at least one XLSX-FIle (test case).\n")

    if "jira_import" in folders:
        pass
    else:
        sys.exit("ERROR: The folder [jira_import] does not exist in the current directory.\n"
                 "Please create a folder in the current directory and start the program again.\n"
                 "NOTE: For safety, this folder should be empty to prevent old results\n"
                 "will be overridden. \n")

    output_folder = "output"
    input_folder = "input"

    print("Programm is starting ...")

    current_dir = os.getcwd()
    input_folder = os.path.join(current_dir, input_folder)

    converter = CreateJiraCsvToImportMultipleTestCasesInXray(input_folder, output_folder)
    converter.convert()

    xlsx_files = os.listdir('input')
    set_test_case_number_to_csv = XlsxSetTestcaseNumber(xlsx_files)
    set_test_case_number_to_csv.process_xlsx_files()

    csv_folder = 'input/'  # Aus diesem ordner werden die xls file genommen
    output_folder = 'jira_import/'  # In diesem ordner wird das neue csv file gespeiechet
    output_csv_name = 'combined.csv'

    processor = ReadXlsxToCSV(
        csv_folder=csv_folder,
        output_folder=output_folder,
        output_csv_name=output_csv_name,
        xlsx_filenames=xlsx_files
    )
    processor.process_csv_files()

    time.sleep(2)

    csv_reader = CSVReader(filename='jira_import/combined.csv',
                           first_column_key="Step Name")

    list_of_lines_to_delete = csv_reader.read_csv()

    first_element = list_of_lines_to_delete.pop(0)

    subtracter = SubtractListElements(list_of_lines_to_delete)
    result = subtracter.subtract_elements()


    time.sleep(2)

    csv_deleter = CSVDeleteLines('jira_import/combined.csv')
    csv_deleter.read_csv()
    csv_deleter.delete_rows(result)
    csv_deleter.save_csv('jira_import/multiple_testcases_file_for_jira.csv')

    time.sleep(2)

    csv_name = "jira_import/multiple_testcases_file_for_jira.csv"
    csv_converter = ChangeCsvCodingFromAsiToUTF8(csv_name)
    csv_converter.convert_encoding()
