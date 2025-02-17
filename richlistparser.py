import csv
import re
from bs4 import BeautifulSoup
import xlsxwriter

# Function to parse HTML file and extract required data
def parse_html_file(input_file, output_file_csv, output_file_excel):
    data_rows = []

    with open(input_file, 'r', encoding='utf-8') as file:
        soup = BeautifulSoup(file, 'html.parser')

        # Find all <a> tags with href containing '/kapital-index/'
        for a_tag in soup.find_all('a', href=re.compile(r"/kapital-index/")):
            print(f"Debug: Found <a> tag: {a_tag}")  # Debugging: Print the <a> tag

            # Navigate through children of the <a> tag
            items = a_tag.find_all('div', class_='c-table__body__row__item')
            if len(items) >= 6:  # Ensure there are at least six divs as described
                name = items[2].find('span').text.strip() if items[2].find('span') else ''
                fortune = items[3].text.strip()
                change = items[4].text.strip()
                tax = items[5].text.strip()
                age = items[6].text.strip()
                _type = items[7].text.strip()

                print(f"Debug: Extracted data - Name: {name}, Fortune: {fortune}, Change: {change}, Tax: {tax}, Age: {age}, Type: {_type}")  # Debugging: Print extracted data

                data_rows.append([name, fortune, change, tax, age, _type])
            else:
                print("Debug: Not enough items found within the <a> tag")  # Debugging: Print when insufficient items are found

    # Write the data to a CSV file
    with open(output_file_csv, 'w', encoding='utf-8', newline='') as csvfile:
        csvwriter = csv.writer(csvfile)
        csvwriter.writerow(["Name", "Fortune", "Change", "Tax", "Age", "Type"])
        csvwriter.writerows(data_rows)

    print(f"Data has been written to {output_file_csv}")

    # Write the data to an Excel file
    workbook = xlsxwriter.Workbook(output_file_excel)
    worksheet = workbook.add_worksheet()

    # Write headers
    headers = ["Name", "Fortune", "Change", "Tax", "Age", "Type"]
    for col_num, header in enumerate(headers):
        worksheet.write(0, col_num, header)

    # Write data rows
    for row_num, row_data in enumerate(data_rows, start=1):
        for col_num, cell_data in enumerate(row_data):
            worksheet.write(row_num, col_num, cell_data)

    workbook.close()
    print(f"Data has been written to {output_file_excel}")

# Replace 'input.html' with your HTML file path, 'richlist.csv' as the output CSV file, and 'richlist.xlsx' as the output Excel file
parse_html_file('richlist.html', 'richlist.csv', 'richlist.xlsx')
