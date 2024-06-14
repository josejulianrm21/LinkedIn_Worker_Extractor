from bs4 import BeautifulSoup
import openpyxl
import os

# Verify the location of the HTML file
file_path = 'data.html'

if not os.path.exists(file_path):
    print(f"The file '{file_path}' does not exist in the specified location.")
    exit()

# Open the HTML file and read its content
with open(file_path, 'r', encoding='utf-8') as file:
    data = file.read()

# Parse the HTML using BeautifulSoup
soup = BeautifulSoup(data, 'html.parser')

# Create a new Excel file and select the first sheet
workbook = openpyxl.Workbook()
sheet = workbook.active
sheet.title = 'Results'

# Write headers in the first row of the Excel file
headers = ['Name', 'Job Title', 'Location', 'Summary']
sheet.append(headers)

# Find all sections containing the desired information
sections = soup.find_all('div', class_='BzflUrAgZIjMzubumWeWNHbvQkiynttqwLmAE')

# Iterate over each section found and extract the data
for section in sections:
    # Exclude sections with the unwanted text
    if 'Contacto de 2.º grado' in section.text:
        continue  # Skip this section and move to the next

    name = section.find('span', class_='entity-result__title-text').text.strip()
    job_title_elem = section.find('div', class_='entity-result__primary-subtitle')
    job_title = job_title_elem.text.strip() if job_title_elem else ''

    location_elem = section.find('div', class_='entity-result__secondary-subtitle')
    location = location_elem.text.strip() if location_elem else ''

    summary_elem = section.find('p', class_='entity-result__summary')
    summary = summary_elem.text.strip() if summary_elem else ''

    # Add the data to a new row in the Excel file
    row = [name, job_title, location, summary]
    sheet.append(row)

# Save the Excel file
output_file = 'results.xlsx'
workbook.save(output_file)

print(f"The results have been saved in '{output_file}'.")
