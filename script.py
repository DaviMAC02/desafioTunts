# Importing CSV module

from csv import reader

# Importing gspread

import gspread


from oauth2client.service_account import ServiceAccountCredentials

from math import ceil


# Defining a main function

def main():
    # Allowing script to access the spreadsheet
    # Note: In order to access the spreadsheet you have to generate your own credentials.json file via Google Cloud Platform

    scope = ["https://spreadsheets.google.com/feeds", 'https://www.googleapis.com/auth/spreadsheets',
            "https://www.googleapis.com/auth/drive.file", "https://www.googleapis.com/auth/drive"]

    credentials = ServiceAccountCredentials.from_json_keyfile_name('credentials.json', scope)
    client = gspread.authorize(credentials)

    # Accessing the CSV file

    try:
        spreadsheet = client.open_by_url("https://docs.google.com/spreadsheets/d/19Fdj2JvkE52SBI-PRyqHQcjvzrk-tsLKJ_tEZfp6CXQ/edit#gid=0")
    except:
        print("Couldn't reach CSV file!")

    # Opening CVS file
    worksheet = spreadsheet.get_worksheet(0)

    # Defining total number of classes

    classes = 50

    # Running through each student

    for student_row in range(4, 28):

        # Getting row values
        student_data = worksheet.row_values(student_row) 
        print('Calculating situation of student number: ' + student_data[0])

        # Calculating student average and getting the number of absences
        average = calculate_average(student_data)
        absences = int(student_data[2])

        # Applying conditions and updating the cells

        if(absences/classes >= 0.25):
            worksheet.update_cell(student_row, 7, 'Reprovado por Falta')
            worksheet.update_cell(student_row, 8, '0')
        elif(average >= 7):
            worksheet.update_cell(student_row, 7, 'Aprovado')
            worksheet.update_cell(student_row, 8, '0')
        elif(average >= 5 and average < 7):
            naf = str(10/average)
            worksheet.update_cell(student_row, 7, 'Exame Final')
            worksheet.update_cell(student_row, 8, naf)


        else:
            worksheet.update_cell(student_row, 7, 'Reprovado por Nota')
            worksheet.update_cell(student_row, 8, '0')

    # Printing a success message

    print("All Done!")




# Defining calculate_average function:

def calculate_average(student_data):
    n1 = float(student_data[3])
    n2 = float(student_data[4])
    n3 = float(student_data[5])

    return ceil((n1 + n2 + n3)/30)


if __name__ == '__main__':
    main()

